//! `Troubleshoot.cmd` — activation/licensing repair menu.
//!
//! Source: `MAS/Separate-Files-Version/Troubleshoot.cmd` (git `f34d025`).
//!
//! The script offers six menu options; each maps to a [`TroubleshootAction`].
//! Only a small amount of the script is *decision* logic — the rest is Windows
//! plumbing (stopping `sppsvc`/`ClipSVC`/`osppsvc`, deleting `tokens.dat`,
//! rebuilding the token stores, running `dism`/`sfc`, rebuilding the WMI
//! repository, repairing Office). That plumbing crosses the [`Spp`] boundary.
//!
//! Portable + unit-tested here:
//!   * The ~31-row HWID (retail digital-license) partial-key list and the
//!     case-insensitive [`is_hwid_key`] matcher (script `:cleanlicensing`).
//!   * The three SPP/OSPP `ApplicationID` GUIDs ([`APPID_WINDOWS`],
//!     [`APPID_OFFICE`], [`APPID_OFFICE_2010`]).
//!   * The `tokens.dat` store path templates ([`SPP_TOKEN_STORES`],
//!     [`OSPP_TOKEN_STORES`]) + [`expand_token_store`], and the valid
//!     `TokenStore` registry values ([`VALID_TOKEN_STORES`]).
//!   * The ClipSVC-rebuild gate [`should_rebuild_clipsvc`] and the per-action
//!     build guards in [`TroubleshootAction::run`].
//!
//! Windows-only (behind the [`Spp`] trait, added via `new_spp_methods`):
//! `dism_restore_health`, `sfc_scannow`, `rebuild_wmi_repository`,
//! `reset_clipsvc_store`, `reset_spp_store`, `reset_ospp_store`, `repair_office`.
//!
//! Deferred (not ported): opening the support web pages for `Help` /
//! `Fix WPA Registry` (a CLI/browser concern), CBS/DISM log compression to the
//! desktop, the SDDL/registry permission-repair (`:fixsppperms`) and the
//! volatile-key ownership dance (`:regown`) — those live inside the coarse
//! `reset_*` effects on the backend.

use crate::error::{Error, Result};
use crate::model::Product;
use crate::platform::{LicenseInfo, Spp};

/// SPP `ApplicationID` for Windows (`_wApp`).
pub const APPID_WINDOWS: &str = "55c92734-d682-4d71-983e-d6ec3f16059f";
/// SPP `ApplicationID` for Office 2013+ / C2R (`_oApp`).
pub const APPID_OFFICE: &str = "0ff1ce15-a989-479d-af46-f275c6370663";
/// OSPP `ApplicationID` for Office 2010 (`_oA14`).
pub const APPID_OFFICE_2010: &str = "59a52881-a989-479d-af46-f275c6370663";

/// The 31 HWID (retail digital-license) partial product keys, verbatim from the
/// script's `:cleanlicensing` match list. These are the last-five-character
/// `PartialProductKey` values Windows reports for a store/digital-license key;
/// their presence means a ClipSVC rebuild is worthwhile.
pub const HWID_PARTIAL_KEYS: &[&str] = &[
    "8HV2C", "QPFCT", "3V66T", "PKCKT", "WXCHW", "8TYMD", "6F4BT", "8HVX7", "KD72Y", "7CFBY",
    "DRR8H", "P39PB", "DYJWX", "MDWWW", "9HKR4", "M7V2X", "2YV77", "WT2RQ", "MHBPB", "QPF8P",
    "2YV66", "VMJ2C", "DJ4F6", "CKFFD", "YY74H", "J8JXD", "BHDCD", "T6R4W", "D32MH", "RRK69",
    "3PJBP",
];

/// Case-insensitive membership test against [`HWID_PARTIAL_KEYS`]
/// (script: `if /i "%_partial%"=="%%#"`).
pub fn is_hwid_key(partial: &str) -> bool {
    let partial = partial.trim();
    HWID_PARTIAL_KEYS
        .iter()
        .any(|k| k.eq_ignore_ascii_case(partial))
}

/// SPP `tokens.dat` store directory templates scanned by the script's `:scandat`
/// (order preserved). Substitute the placeholders with [`expand_token_store`].
pub const SPP_TOKEN_STORES: &[&str] = &[
    r"%SysPath%\spp\store_test\2.0\",
    r"%SysPath%\spp\store\",
    r"%SysPath%\spp\store\2.0\",
    r"%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform\",
    r"%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareLicensing\",
];

/// OSPP `tokens.dat` store directory template (`:scandatospp`).
pub const OSPP_TOKEN_STORES: &[&str] = &[r"%ProgramData%\Microsoft\OfficeSoftwareProtectionPlatform\"];

/// Values the `TokenStore` registry entry is allowed to hold (`%SysPath%`
/// still to be expanded); anything else the script flags as a bad registry.
pub const VALID_TOKEN_STORES: &[&str] = &[
    r"%SysPath%\spp\store",
    r"%SysPath%\spp\store\2.0",
    r"%SysPath%\spp\store_test\2.0",
];

/// Expand a store template's `%SysPath%` / `%Systemdrive%` / `%ProgramData%`
/// placeholders. Pure so the backend and tests share one substitution.
pub fn expand_token_store(template: &str, sys_path: &str, system_drive: &str, program_data: &str) -> String {
    template
        .replace("%SysPath%", sys_path)
        .replace("%Systemdrive%", system_drive)
        .replace("%ProgramData%", program_data)
}

/// Whether the ClipSVC license rebuild sub-flow should run (script
/// `:cleanlicensing` gate): Windows 10+ (`>= 10240`), not already permanently
/// activated, and an HWID/digital-license key is installed. The internet /
/// licensing-server reachability checks that follow are Windows effects and
/// live inside [`Spp::reset_clipsvc_store`].
pub fn should_rebuild_clipsvc(build: u32, windows_products: &[LicenseInfo]) -> bool {
    if build < 10240 {
        return false;
    }
    // Permanently activated = Licensed with no grace remaining (OEM/retail/
    // digital); the script skips the rebuild in that case.
    let permanent = windows_products
        .iter()
        .any(|p| p.is_activated() && p.grace_minutes == Some(0));
    if permanent {
        return false;
    }
    windows_products
        .iter()
        .any(|p| p.partial_product_key.as_deref().is_some_and(is_hwid_key))
}

/// The Troubleshoot menu options (`:at_menu`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TroubleshootAction {
    /// `[1]` Open the online help / troubleshooting pages (informational).
    Help,
    /// `[2]` `dism /english /online /cleanup-image /restorehealth`.
    DismRestoreHealth,
    /// `[3]` `sfc /scannow`.
    SfcScannow,
    /// `[4]` Rebuild the WMI repository.
    FixWmi,
    /// `[5]` Fix Licensing (ClipSVC + SPP + OSPP + Office repair).
    FixLicensing,
    /// `[6]` Open the WPA-registry-fix page (informational).
    FixWpaRegistry,
}

impl TroubleshootAction {
    /// Menu order, matching the script.
    pub const ALL: [TroubleshootAction; 6] = [
        TroubleshootAction::Help,
        TroubleshootAction::DismRestoreHealth,
        TroubleshootAction::SfcScannow,
        TroubleshootAction::FixWmi,
        TroubleshootAction::FixLicensing,
        TroubleshootAction::FixWpaRegistry,
    ];

    pub const fn label(self) -> &'static str {
        match self {
            TroubleshootAction::Help => "Help",
            TroubleshootAction::DismRestoreHealth => "Dism RestoreHealth",
            TroubleshootAction::SfcScannow => "SFC Scannow",
            TroubleshootAction::FixWmi => "Fix WMI",
            TroubleshootAction::FixLicensing => "Fix Licensing",
            TroubleshootAction::FixWpaRegistry => "Fix WPA Registry",
        }
    }

    /// Whether the action only opens a support web page (no local effect).
    pub const fn is_informational(self) -> bool {
        matches!(self, TroubleshootAction::Help | TroubleshootAction::FixWpaRegistry)
    }

    /// Run the action. Requires elevation; applies the script's build guards
    /// before delegating each effect across the [`Spp`] boundary.
    pub fn run(self, spp: &dyn Spp) -> Result<()> {
        if self.is_informational() {
            // Help / Fix WPA Registry only launch a browser; the CLI handles
            // that, the core has nothing to do.
            return Ok(());
        }
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        match self {
            TroubleshootAction::DismRestoreHealth => {
                if spp.windows_build()? < 9200 {
                    return Err(unsupported("DISM RestoreHealth requires Windows 8/8.1/10/11 or Server equivalents"));
                }
                spp.dism_restore_health()
            }
            TroubleshootAction::SfcScannow => spp.sfc_scannow(),
            TroubleshootAction::FixWmi => spp.rebuild_wmi_repository(),
            TroubleshootAction::FixLicensing => fix_licensing(spp),
            TroubleshootAction::Help | TroubleshootAction::FixWpaRegistry => unreachable!(),
        }
    }
}

/// Fix Licensing flow (`:retokens`): optional ClipSVC rebuild (gated), then
/// SPP + OSPP token-store rebuilds, then the Office repair trigger. The gate is
/// portable; each rebuild is a coarse Windows effect.
fn fix_licensing(spp: &dyn Spp) -> Result<()> {
    let build = spp.windows_build()?;
    if build == 6001 {
        return Err(unsupported("Fix Licensing is not supported on Windows Vista SP1; upgrade to SP2"));
    }
    if should_rebuild_clipsvc(build, &spp.installed_products(Product::Windows)?) {
        spp.reset_clipsvc_store()?;
    }
    spp.reset_spp_store()?;
    spp.reset_ospp_store()?;
    spp.repair_office()?;
    Ok(())
}

fn unsupported(what: &str) -> Error {
    Error::Unsupported { what: what.into() }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::LicenseStatus;
    use crate::platform::test_util::FakeSpp;

    fn win(key: Option<&str>, activated: bool, grace: Option<u32>) -> LicenseInfo {
        LicenseInfo {
            activation_id: "win".into(),
            name: "Windows".into(),
            description: String::new(),
            partial_product_key: key.map(str::to_string),
            status: LicenseStatus::from_wmi(if activated { 1 } else { 0 }).unwrap(),
            license_family: None,
            grace_minutes: grace,
        }
    }

    #[test]
    fn hwid_list_is_verbatim() {
        assert_eq!(HWID_PARTIAL_KEYS.len(), 31);
        assert_eq!(HWID_PARTIAL_KEYS[0], "8HV2C");
        assert_eq!(HWID_PARTIAL_KEYS[30], "3PJBP");
    }

    #[test]
    fn is_hwid_key_matches_case_insensitively() {
        assert!(is_hwid_key("8HV2C"));
        assert!(is_hwid_key("8hv2c")); // case-insensitive
        assert!(is_hwid_key("  3PJBP  ")); // trimmed
        assert!(!is_hwid_key("ABCDE"));
        assert!(!is_hwid_key(""));
    }

    #[test]
    fn appid_guids_are_verbatim() {
        assert_eq!(APPID_WINDOWS, "55c92734-d682-4d71-983e-d6ec3f16059f");
        assert_eq!(APPID_OFFICE, "0ff1ce15-a989-479d-af46-f275c6370663");
        assert_eq!(APPID_OFFICE_2010, "59a52881-a989-479d-af46-f275c6370663");
        // Windows AppID must agree with the domain model's constant.
        assert_eq!(APPID_WINDOWS, Product::Windows.application_id());
    }

    #[test]
    fn expand_token_store_substitutes_placeholders() {
        let got = expand_token_store(
            r"%SysPath%\spp\store\",
            r"C:\Windows\System32",
            "C:",
            r"C:\ProgramData",
        );
        assert_eq!(got, r"C:\Windows\System32\spp\store\");
        let ospp = expand_token_store(OSPP_TOKEN_STORES[0], "sp", "sd", "PD");
        assert_eq!(ospp, r"PD\Microsoft\OfficeSoftwareProtectionPlatform\");
    }

    #[test]
    fn clipsvc_gate() {
        // Too old.
        assert!(!should_rebuild_clipsvc(7601, &[win(Some("8HV2C"), false, None)]));
        // HWID key present, not permanently activated -> rebuild.
        assert!(should_rebuild_clipsvc(19045, &[win(Some("8HV2C"), false, None)]));
        // Permanently activated (Licensed, grace 0) -> skip.
        assert!(!should_rebuild_clipsvc(19045, &[win(Some("8HV2C"), true, Some(0))]));
        // No HWID key installed -> skip.
        assert!(!should_rebuild_clipsvc(19045, &[win(Some("XXXXX"), false, None)]));
        // No products at all -> skip.
        assert!(!should_rebuild_clipsvc(19045, &[]));
    }

    #[test]
    fn action_menu_shape() {
        assert_eq!(TroubleshootAction::ALL.len(), 6);
        assert_eq!(TroubleshootAction::DismRestoreHealth.label(), "Dism RestoreHealth");
        assert!(TroubleshootAction::Help.is_informational());
        assert!(!TroubleshootAction::FixLicensing.is_informational());
    }

    #[test]
    fn informational_actions_are_noops() {
        let spp = FakeSpp::default();
        TroubleshootAction::Help.run(&spp).unwrap();
        TroubleshootAction::FixWpaRegistry.run(&spp).unwrap();
    }

    #[test]
    fn effect_actions_require_elevation() {
        let spp = FakeSpp { elevated: false, ..Default::default() };
        assert!(matches!(TroubleshootAction::SfcScannow.run(&spp), Err(Error::NotElevated)));
    }

    #[test]
    fn dism_guard_rejects_old_builds() {
        let spp = FakeSpp { build: 7601, ..Default::default() };
        assert!(matches!(
            TroubleshootAction::DismRestoreHealth.run(&spp),
            Err(Error::Unsupported { .. })
        ));
    }

    #[test]
    fn fix_licensing_rejects_vista_sp1() {
        let spp = FakeSpp { build: 6001, ..Default::default() };
        assert!(matches!(
            TroubleshootAction::FixLicensing.run(&spp),
            Err(Error::Unsupported { .. })
        ));
    }
}
