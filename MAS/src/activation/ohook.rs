//! `Ohook_Activation_AIO.cmd` — permanent Office activation by hooking the SPP.
//!
//! Ohook does not talk to any server. It drops a reverse-engineered SPP-client
//! DLL (`sppc32.dll` / `sppc64.dll`, standing in for `OSPPC.DLL`) that forwards
//! to the renamed genuine `sppcs.dll` while overriding the license-status calls
//! so Office always reports "licensed". Those DLLs are opaque binary blobs with
//! **pinned SHA-256 hashes**; a faithful port ships the identical bytes — Rust
//! cannot reimplement the SPP internals, so they stay **external assets**.
//!
//! ## What is ported here (pure, unit-tested)
//! * the two embedded product tables ([`OhookEntry`] / [`MsiEntry`]) parsed
//!   from [`data`] — verbatim transcription, `%f%` stripped from keys;
//! * table selection ([`lookup`]) — `(version, edition) -> key / activation id`;
//! * build→WMI-class selection ([`spp_class`], `:oh_setspp`) — `>=9200 ⇒
//!   SoftwareLicensing*`, else / Office-2010 ⇒ `OfficeSoftwareProtection*`;
//! * the license-file base-name transforms ([`derive_license_name`],
//!   `:oh_installlic` fallback) and the whole install [`plan_install`] —
//!   which `.xrm-ms` prefixes to install and where the DLL is dropped for
//!   Click-to-Run (`<root>\vfs\System[X86]`) vs MSI (the InstallRoot itself).
//!
//! ## Windows-only (through the [`Spp`] trait — see `new_spp_methods`)
//! Office-install discovery, the pinned-hash blob presence check, the
//! `.xrm-ms` glob, and the DLL rename/symlink dance (`:oh_hookinstall*`).
//! The licenses themselves go in via the existing [`Spp::install_license`].
//!
//! ## Deferred
//! Uninstall (`:oh_uninstall`), the `integrator.exe` fast path, the anti-banner
//! registry keys, and the per-install DLL timestamp reskin (`:oh_extractdll`,
//! offsets 2564/3076) — none are needed to activate and all are pure Windows
//! side effects. `run` refuses honestly until the blob assets are present.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;
use std::path::{Path, PathBuf};

#[path = "ohook_data.rs"]
mod data;

/// One row of `:ohookdata` — a Click-to-Run / VL product and its GVLK/MAK key.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OhookEntry {
    /// Office major version: 14, 15 or 16.
    pub ver: u8,
    /// SPP activation-id GUID.
    pub act_id: String,
    /// Product key with the `%f%` splitter removed.
    pub key: String,
    /// License channel token (`Retail`, `MAK`, `Subscription`, …).
    pub lic_suffix: String,
    /// Edition id matched against the detected product (`ProPlus2021Volume`).
    pub edition: String,
    /// Reference-only sibling edition ids (`[HomeBusinessDemoR]`), if any.
    pub other_ids: Option<String>,
}

/// One row of `:msiofficedata` — maps an MSI product-code hex to an edition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MsiEntry {
    pub ver: u8,
    pub act_id: String,
    /// 4-hex product id embedded in the MSI ProductCode (`-0015-`).
    pub product_id_hex: String,
    pub edition: String,
}

/// Split on `_`, collapsing runs of `_` exactly as `for /f delims=_` does
/// (the script pads columns with `____` for alignment).
fn fields(line: &str) -> Vec<&str> {
    line.split('_').filter(|s| !s.is_empty()).collect()
}

/// Parse [`data::OHOOK_TABLE`]. Malformed / comment (`::`) lines are skipped.
pub fn parse_ohook_table(raw: &str) -> Vec<OhookEntry> {
    raw.lines()
        .map(str::trim)
        .filter(|l| !l.is_empty() && !l.starts_with("::"))
        .filter_map(|l| {
            let f = fields(l);
            if f.len() < 5 {
                return None;
            }
            Some(OhookEntry {
                ver: f[0].parse().ok()?,
                act_id: f[1].to_string(),
                key: f[2].replace("%f%", ""),
                lic_suffix: f[3].to_string(),
                edition: f[4].to_string(),
                other_ids: if f.len() > 5 {
                    Some(f[5..].join("_"))
                } else {
                    None
                },
            })
        })
        .collect()
}

/// Parse [`data::MSI_TABLE`].
pub fn parse_msi_table(raw: &str) -> Vec<MsiEntry> {
    raw.lines()
        .map(str::trim)
        .filter(|l| !l.is_empty() && !l.starts_with("::"))
        .filter_map(|l| {
            let f = fields(l);
            if f.len() < 4 {
                return None;
            }
            Some(MsiEntry {
                ver: f[0].parse().ok()?,
                act_id: f[1].to_string(),
                product_id_hex: f[2].to_string(),
                edition: f[3].to_string(),
            })
        })
        .collect()
}

/// The parsed `:ohookdata` table.
pub fn ohook_table() -> Vec<OhookEntry> {
    parse_ohook_table(data::OHOOK_TABLE)
}

/// The parsed `:msiofficedata` table.
pub fn msi_table() -> Vec<MsiEntry> {
    parse_msi_table(data::MSI_TABLE)
}

/// `:ohookdata getinfo` — first row matching `(ver, edition)`, case-insensitive.
pub fn lookup<'a>(table: &'a [OhookEntry], ver: u8, edition: &str) -> Option<&'a OhookEntry> {
    table
        .iter()
        .find(|e| e.ver == ver && e.edition.eq_ignore_ascii_case(edition))
}

/// `:msiofficedata` — row for a `(ver, productHex)` pair, case-insensitive.
pub fn msi_edition<'a>(table: &'a [MsiEntry], ver: u8, hex: &str) -> Option<&'a MsiEntry> {
    table
        .iter()
        .find(|e| e.ver == ver && e.product_id_hex.eq_ignore_ascii_case(hex))
}

/// Processor architecture of the Office install (`_oArch`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Arch {
    X86,
    X64,
}

/// Install layout of the Office suite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfficeKind {
    /// Click-to-Run (Office 15/16): DLL lands in `<root>\vfs\System[X86]`.
    C2R,
    /// Windows Installer (Office 14/15/16): DLL lands in the InstallRoot.
    Msi,
}

/// Which SPP the product talks to — decides the WMI classes and the DLL dance.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SppClass {
    /// `SoftwareLicensingProduct` / `SoftwareLicensingService` (Windows 8+).
    SoftwareLicensing,
    /// `OfficeSoftwareProtection*` (Office 2010, or Windows 7 and older).
    Ospp,
}

impl SppClass {
    pub const fn product_class(self) -> &'static str {
        match self {
            SppClass::SoftwareLicensing => "SoftwareLicensingProduct",
            SppClass::Ospp => "OfficeSoftwareProtectionProduct",
        }
    }

    pub const fn service_class(self) -> &'static str {
        match self {
            SppClass::SoftwareLicensing => "SoftwareLicensingService",
            SppClass::Ospp => "OfficeSoftwareProtectionService",
        }
    }

    /// `:oh_hookinstall_ospp` is used for the OSPP dance, `:oh_hookinstall`
    /// otherwise.
    pub const fn is_ospp(self) -> bool {
        matches!(self, SppClass::Ospp)
    }
}

/// `:oh_setspp` — OSPP for Office 2010 (`ver 14`) or Windows < 8 (build 9200);
/// `SoftwareLicensing*` otherwise.
pub fn spp_class(windows_build: u32, office_ver: u8) -> SppClass {
    if office_ver == 14 || windows_build < 9200 {
        SppClass::Ospp
    } else {
        SppClass::SoftwareLicensing
    }
}

/// Custom SPP-client DLL file name for the install's architecture.
pub const fn dll_name(arch: Arch) -> &'static str {
    match arch {
        Arch::X64 => "sppc64.dll",
        Arch::X86 => "sppc32.dll",
    }
}

/// Where the custom DLL is dropped (`_hookPath`). C2R: `<root>\vfs\System[X86]`;
/// MSI: the InstallRoot itself.
pub fn hook_dir(kind: OfficeKind, install_root: &Path, arch: Arch) -> PathBuf {
    match kind {
        OfficeKind::Msi => install_root.to_path_buf(),
        OfficeKind::C2R => install_root.join("vfs").join(match arch {
            Arch::X64 => "System",
            Arch::X86 => "SystemX86",
        }),
    }
}

/// C2R license directory (`_oLPath`): `Licenses16` for v16, `Licenses` for v15.
pub fn license_dir(install_root: &Path, office_ver: u8) -> PathBuf {
    install_root.join(if office_ver >= 16 {
        "Licenses16"
    } else {
        "Licenses"
    })
}

/// `:oh_installlic` fallback name transforms: edition id → `.xrm-ms` base name.
/// Order is significant and mirrors the script's `set _License=%_License:a=b%`
/// chain (each replaces *all* occurrences).
pub fn derive_license_name(edition: &str, preview: bool) -> String {
    let mut s = edition.to_string();
    s = s.replace("XVolume", "XC2RVL_");
    s = s.replace("O365EduCloudRetail", "O365EduCloudEDUR_");
    s = s.replace("ProjectProRetail", "ProjectProO365R_");
    s = s.replace("ProjectStdRetail", "ProjectStdO365R_");
    s = s.replace("VisioProRetail", "VisioProO365R_");
    s = s.replace("VisioStdRetail", "VisioStdO365R_");
    if preview {
        s = s.replace("Volume", "PreviewVL_");
    }
    s = s.replace("Retail", "R_");
    s = s.replace("Volume", "VL_");
    s
}

/// A discovered Office installation (built on Windows by `ohook_office_installs`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OfficeInstall {
    pub kind: OfficeKind,
    pub office_ver: u8,
    pub arch: Arch,
    /// `_oRoot`: the `...\root` folder for C2R, the InstallRoot for MSI.
    pub root: PathBuf,
    /// `_oIds`: detected product-edition ids (may carry a `-Preview` suffix).
    pub editions: Vec<String>,
}

/// One resolved product: its key/activation-id from the table plus the derived
/// C2R license base name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ResolvedProduct {
    pub edition: String,
    pub act_id: String,
    pub key: String,
    pub lic_suffix: String,
    pub license_name: String,
}

/// Where and how the DLL hook is placed (consumed by `ohook_place_hook`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HookPlan {
    pub kind: OfficeKind,
    pub office_ver: u8,
    pub arch: Arch,
    pub hook_dir: PathBuf,
    pub dll_name: &'static str,
    pub class: SppClass,
}

/// `.xrm-ms` files to (re)install for a C2R suite.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LicensePlan {
    pub dir: PathBuf,
    /// Glob prefixes; each is installed as `<prefix>*.xrm-ms`.
    pub prefixes: Vec<String>,
}

/// The full, purely-computed plan for one Office install.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct InstallPlan {
    pub hook: HookPlan,
    /// `None` for MSI: it reuses licenses already on disk (`_oLPath` unset).
    pub license: Option<LicensePlan>,
    pub products: Vec<ResolvedProduct>,
}

/// Assemble the install plan for `install` from the parsed `table`.
pub fn plan_install(
    install: &OfficeInstall,
    windows_build: u32,
    table: &[OhookEntry],
) -> InstallPlan {
    let class = spp_class(windows_build, install.office_ver);

    let hook = HookPlan {
        kind: install.kind,
        office_ver: install.office_ver,
        arch: install.arch,
        hook_dir: hook_dir(install.kind, &install.root, install.arch),
        dll_name: dll_name(install.arch),
        class,
    };

    let mut products = Vec::new();
    for ed in &install.editions {
        // Preview builds carry a `-Preview` edition suffix in the table and in
        // `_prod`, but the license base name is derived from the bare edition.
        let preview = ed.ends_with("-Preview");
        let base = ed.trim_end_matches("-Preview");
        if let Some(e) = lookup(table, install.office_ver, ed) {
            products.push(ResolvedProduct {
                edition: ed.clone(),
                act_id: e.act_id.clone(),
                key: e.key.clone(),
                lic_suffix: e.lic_suffix.clone(),
                license_name: derive_license_name(base, preview),
            });
        }
    }

    let license = match install.kind {
        OfficeKind::C2R => {
            let mut prefixes = vec!["client-issuance-".to_string()];
            prefixes.extend(products.iter().map(|p| p.license_name.clone()));
            prefixes.push("pkeyconfig-office".to_string());
            Some(LicensePlan {
                dir: license_dir(&install.root, install.office_ver),
                prefixes,
            })
        }
        OfficeKind::Msi => None,
    };

    InstallPlan {
        hook,
        license,
        products,
    }
}

pub struct Ohook;

impl Activator for Ohook {
    fn method(&self) -> Method {
        Method::Ohook
    }

    fn run(&self, spp: &dyn Spp, product: Product) -> Result<()> {
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        if product != Product::Office {
            return Err(Error::Unsupported {
                what: "Ohook only activates Office".into(),
            });
        }

        // The pinned-SHA-256 SPP-client blobs are mandatory and never embedded.
        // Absent (or non-Windows) → refuse with an actionable message rather
        // than half-installing.
        if !matches!(spp.ohook_assets_present(), Ok(true)) {
            return Err(Error::Unsupported {
                what: "Ohook needs the pinned SPP-client DLL assets \
                       (sppc32.dll / sppc64.dll); ship them alongside MAS and retry"
                    .into(),
            });
        }

        let installs = spp.ohook_office_installs()?;
        if installs.is_empty() {
            return Err(Error::Unsupported {
                what: "no supported Office install found (C2R 15/16 or MSI 14/15/16)".into(),
            });
        }

        let build = spp.windows_build()?;
        let table = ohook_table();
        for install in &installs {
            let plan = plan_install(install, build, &table);

            // C2R: (re)install the license files the plan selected. MSI reuses
            // the licenses already present, so `license` is None there.
            if let Some(lic) = &plan.license {
                for prefix in &lic.prefixes {
                    for file in spp.ohook_glob_licenses(&lic.dir, prefix)? {
                        spp.install_license(&file)?;
                    }
                }
            }

            spp.ohook_place_hook(&plan.hook)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::platform::test_util::FakeSpp;

    #[test]
    fn tables_parse_to_expected_counts() {
        let oh = ohook_table();
        let msi = msi_table();
        assert_eq!(oh.len(), 235, "ohookdata row count");
        assert_eq!(msi.len(), 148, "msiofficedata row count");
        assert_eq!(oh.iter().filter(|e| e.ver == 14).count(), 32);
        assert_eq!(oh.iter().filter(|e| e.ver == 15).count(), 52);
        assert_eq!(oh.iter().filter(|e| e.ver == 16).count(), 151);
        assert_eq!(msi.iter().filter(|e| e.ver == 14).count(), 43);
    }

    #[test]
    fn key_strips_the_f_splitter() {
        let oh = ohook_table();
        // 2010 Access Retail: 7KTYC-XR43P-C3MRW-BJKFD-XB%f%YPG.
        let e = lookup(&oh, 14, "AccessR").unwrap();
        assert_eq!(e.key, "7KTYC-XR43P-C3MRW-BJKFD-XBYPG");
        assert!(!e.key.contains("%f%"));
        assert_eq!(e.act_id, "4d463c2c-0505-4626-8cdb-a4da82e2d8ed");
        assert_eq!(e.lic_suffix, "Retail");
    }

    #[test]
    fn every_key_is_clean_and_five_groups() {
        for e in ohook_table() {
            assert!(!e.key.contains("%f%"), "{}: {}", e.edition, e.key);
            assert_eq!(
                e.key.split('-').count(),
                5,
                "{} key not 5 groups: {}",
                e.edition,
                e.key
            );
        }
    }

    #[test]
    fn lookup_is_version_scoped_and_case_insensitive() {
        let oh = ohook_table();
        // Same edition string, different Office versions → different rows.
        let p15 = lookup(&oh, 15, "ProPlusVolume").unwrap();
        let p16 = lookup(&oh, 16, "ProPlusVolume").unwrap();
        assert_ne!(p15.key, p16.key);
        assert!(lookup(&oh, 16, "proplus2021volume").is_some());
        assert!(lookup(&oh, 99, "ProPlusVolume").is_none());
    }

    #[test]
    fn other_ids_captured_for_reference_rows() {
        let oh = ohook_table();
        let hb = lookup(&oh, 14, "HomeBusinessR").unwrap();
        assert_eq!(hb.other_ids.as_deref(), Some("[HomeBusinessDemoR]"));
        let ap = lookup(&oh, 14, "AccessR").unwrap();
        assert_eq!(ap.other_ids, None);
    }

    #[test]
    fn msi_hex_maps_to_edition() {
        let msi = msi_table();
        let e = msi_edition(&msi, 14, "0015").unwrap();
        assert_eq!(e.edition, "AccessR");
        assert!(msi_edition(&msi, 16, "001B").is_some()); // WordRetail/Volume
    }

    #[test]
    fn spp_class_selection() {
        // Windows 10, modern Office → SoftwareLicensing.
        assert_eq!(spp_class(19045, 16), SppClass::SoftwareLicensing);
        // Office 2010 is always OSPP regardless of OS.
        assert_eq!(spp_class(19045, 14), SppClass::Ospp);
        // Windows 7 (build < 9200) → OSPP.
        assert_eq!(spp_class(7601, 16), SppClass::Ospp);
        assert_eq!(
            SppClass::SoftwareLicensing.product_class(),
            "SoftwareLicensingProduct"
        );
        assert_eq!(
            SppClass::Ospp.service_class(),
            "OfficeSoftwareProtectionService"
        );
    }

    #[test]
    fn dll_and_hook_paths() {
        assert_eq!(dll_name(Arch::X64), "sppc64.dll");
        assert_eq!(dll_name(Arch::X86), "sppc32.dll");

        let root = Path::new(r"C:\Program Files\Microsoft Office\root");
        assert_eq!(
            hook_dir(OfficeKind::C2R, root, Arch::X64),
            root.join("vfs").join("System")
        );
        assert_eq!(
            hook_dir(OfficeKind::C2R, root, Arch::X86),
            root.join("vfs").join("SystemX86")
        );
        // MSI drops the DLL straight into the InstallRoot.
        let msi_root = Path::new(r"C:\Program Files\Microsoft Office\Office16");
        assert_eq!(hook_dir(OfficeKind::Msi, msi_root, Arch::X64), msi_root);

        assert_eq!(license_dir(root, 16), root.join("Licenses16"));
        assert_eq!(license_dir(root, 15), root.join("Licenses"));
    }

    #[test]
    fn license_name_transforms() {
        // Plain volume / retail.
        assert_eq!(
            derive_license_name("ProPlus2021Volume", false),
            "ProPlus2021VL_"
        );
        assert_eq!(derive_license_name("AccessRetail", false), "AccessR_");
        // Preview replaces Volume before the generic pass.
        assert_eq!(
            derive_license_name("ProPlus2024Volume", true),
            "ProPlus2024PreviewVL_"
        );
        // Specific-before-generic ordering: ProjectProRetail must not become
        // "ProjectProR_".
        assert_eq!(
            derive_license_name("ProjectProRetail", false),
            "ProjectProO365R_"
        );
        // MAKC2R "XVolume".
        assert_eq!(
            derive_license_name("ProjectProXVolume", false),
            "ProjectProXC2RVL_"
        );
    }

    #[test]
    fn c2r_plan_selects_licenses_and_vfs_dll() {
        let table = ohook_table();
        let install = OfficeInstall {
            kind: OfficeKind::C2R,
            office_ver: 16,
            arch: Arch::X64,
            root: PathBuf::from(r"C:\Program Files\Microsoft Office\root"),
            editions: vec!["ProPlus2021Volume".into()],
        };
        let plan = plan_install(&install, 19045, &table);

        assert_eq!(plan.hook.dll_name, "sppc64.dll");
        assert_eq!(plan.hook.class, SppClass::SoftwareLicensing);
        assert!(plan.hook.hook_dir.ends_with("System"));

        let lic = plan.license.expect("C2R installs licenses");
        assert_eq!(lic.prefixes.first().unwrap(), "client-issuance-");
        assert!(lic.prefixes.iter().any(|p| p == "ProPlus2021VL_"));
        assert_eq!(lic.prefixes.last().unwrap(), "pkeyconfig-office");

        assert_eq!(plan.products.len(), 1);
        assert_eq!(plan.products[0].lic_suffix, "MAK-AE1");
    }

    #[test]
    fn msi_plan_has_no_license_install() {
        let table = ohook_table();
        let install = OfficeInstall {
            kind: OfficeKind::Msi,
            office_ver: 14,
            arch: Arch::X86,
            root: PathBuf::from(r"C:\Program Files\Microsoft Office\Office14"),
            editions: vec!["ProPlusVL".into()],
        };
        let plan = plan_install(&install, 7601, &table);
        assert!(plan.license.is_none());
        assert_eq!(plan.hook.class, SppClass::Ospp); // ver 14 ⇒ OSPP
        assert_eq!(plan.hook.dll_name, "sppc32.dll");
        assert_eq!(plan.hook.hook_dir, install.root);
    }

    #[test]
    fn preview_edition_resolves_preview_row_and_name() {
        let table = ohook_table();
        let install = OfficeInstall {
            kind: OfficeKind::C2R,
            office_ver: 16,
            arch: Arch::X64,
            root: PathBuf::from(r"C:\Office\root"),
            editions: vec!["ProPlus2024Volume-Preview".into()],
        };
        let plan = plan_install(&install, 19045, &table);
        assert_eq!(plan.products.len(), 1, "matched the -Preview table row");
        let lic = plan.license.unwrap();
        assert!(lic.prefixes.iter().any(|p| p == "ProPlus2024PreviewVL_"));
    }

    #[test]
    fn run_requires_elevation() {
        let spp = FakeSpp {
            elevated: false,
            ..Default::default()
        };
        assert!(matches!(
            Ohook.run(&spp, Product::Office),
            Err(Error::NotElevated)
        ));
    }

    #[test]
    fn run_rejects_non_office() {
        let spp = FakeSpp::default();
        assert!(matches!(
            Ohook.run(&spp, Product::Windows),
            Err(Error::Unsupported { .. })
        ));
    }

    #[test]
    fn run_refuses_without_blob_assets() {
        // Default fake leaves the ohook_* methods at their UnsupportedPlatform
        // default → assets check fails → honest Unsupported naming the DLLs.
        let spp = FakeSpp::default();
        let err = Ohook.run(&spp, Product::Office).unwrap_err();
        match err {
            Error::Unsupported { what } => assert!(what.contains("SPP-client DLL")),
            other => panic!("expected Unsupported, got {other:?}"),
        }
    }
}
