//! The Software Protection Platform (SPP) boundary.
//!
//! This is the single seam between portable logic and Windows internals. Every
//! privileged thing the scripts do — WMI calls against `SoftwareLicensingService`,
//! registry edits, spawning `ClipUp.exe`/`slmgr.vbs` — is expressed here as a
//! safe trait. The rest of the crate programs against [`Spp`] and never touches
//! `unsafe` or the `windows` crate directly (constraint 3: encapsulate FFI behind
//! a sound safe API).
//!
//! Two implementations exist, chosen at compile time:
//!
//! * [`windows_backend`] (`--features winapi`, Windows target) — the real thing.
//! * [`stub`] (everywhere else) — read-only calls return empty/best-effort data
//!   and every privileged call returns [`Error::UnsupportedPlatform`], so the
//!   portable core builds and tests on Linux.

use crate::error::{Error, Result};
use crate::model::{LicenseStatus, Product};

#[cfg(all(windows, feature = "winapi"))]
pub mod windows_backend;
#[cfg(not(all(windows, feature = "winapi")))]
pub mod stub;
#[cfg(test)]
pub mod test_util;

/// One product entry as reported by the SPP (`SoftwareLicensingProduct` row).
#[derive(Debug, Clone)]
pub struct LicenseInfo {
    /// Activation ID (the WMI object `ID`), used to target `Activate()` etc.
    pub activation_id: String,
    /// Human-readable product name, e.g. `Windows(R), Professional edition`.
    pub name: String,
    /// SPP `Description`, e.g. `Windows(R) Operating System, RETAIL channel`.
    pub description: String,
    /// Last five characters of the installed key, or `None` if no key.
    pub partial_product_key: Option<String>,
    /// Current license state.
    pub status: LicenseStatus,
    /// `LicenseFamily`, used to distinguish Office apps (Project/Visio/…).
    pub license_family: Option<String>,
    /// Seconds of grace remaining (KMS / trial), if applicable.
    pub grace_minutes: Option<u32>,
}

impl LicenseInfo {
    pub fn is_activated(&self) -> bool {
        self.status.is_activated()
    }
}

/// Safe abstraction over the Windows Software Protection Platform.
///
/// Method names mirror the underlying WMI methods so the mapping back to the
/// original scripts stays obvious.
pub trait Spp {
    /// Query all licensable products of a family
    /// (`SELECT … FROM SoftwareLicensingProduct WHERE ApplicationID='…' AND
    /// PartialProductKey IS NOT NULL`).
    fn installed_products(&self, product: Product) -> Result<Vec<LicenseInfo>>;

    /// `SoftwareLicensingService.InstallProductKey(key)` — install a GVLK/retail key.
    fn install_product_key(&self, product: Product, key: &str) -> Result<()>;

    /// `SoftwareLicensingProduct.UninstallProductKey()` for the given activation ID.
    fn uninstall_product_key(&self, product: Product, activation_id: &str) -> Result<()>;

    /// `SetKeyManagementServiceMachine` + `SetKeyManagementServicePort`.
    fn set_kms_host(&self, product: Product, host: &str, port: u16) -> Result<()>;

    /// `ClearKeyManagementServiceMachine` — revert to auto-discovery.
    fn clear_kms_host(&self, product: Product) -> Result<()>;

    /// `SoftwareLicensingProduct.Activate()` for the given activation ID.
    fn activate(&self, product: Product, activation_id: &str) -> Result<()>;

    /// `SoftwareLicensingService.InstallLicense(<xml>)` — install a license/token file.
    fn install_license(&self, xml_path: &std::path::Path) -> Result<()>;

    /// OS build number (e.g. `19045`). Drives method eligibility.
    fn windows_build(&self) -> Result<u32>;

    /// Current Windows edition ID (e.g. `Professional`).
    fn windows_edition(&self) -> Result<String>;

    /// Whether the current process is elevated (administrator).
    fn is_elevated(&self) -> bool;

    // --- Digital-license / ClipSVC operations (HWID, KMS38) ------------------
    // These drive the closed ClipSVC path: MAS itself only generates a ticket
    // and hands it to Windows, so the port does the same behind this boundary.

    /// Produce a `GenuineTicket.xml` for this machine (hardware hash + region),
    /// returning its bytes. Windows-only (uses the SPP/ClipUp machinery).
    fn generate_genuine_ticket(&self) -> Result<Vec<u8>> {
        Err(Error::UnsupportedPlatform {
            operation: "generate genuine ticket",
        })
    }

    /// Write a genuine ticket to the ClipSVC drop path
    /// (`%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket\GenuineTicket.xml`).
    fn write_genuine_ticket(&self, _xml: &[u8]) -> Result<()> {
        Err(Error::UnsupportedPlatform {
            operation: "write genuine ticket",
        })
    }

    /// Restart a Windows service by name (e.g. `ClipSVC`).
    fn restart_service(&self, _name: &str) -> Result<()> {
        Err(Error::UnsupportedPlatform {
            operation: "restart service",
        })
    }

    /// Run `ClipUp.exe` with the given args (e.g. `-v -o` to apply a ticket).
    fn run_clipup(&self, _args: &[&str]) -> Result<()> {
        Err(Error::UnsupportedPlatform {
            operation: "run ClipUp",
        })
    }

    /// Whether ClipSVC has produced `tokens.dat` — the HWID success check.
    fn clip_tokens_present(&self) -> bool {
        false
    }

    /// Pin one Windows product's KMS host to loopback (`127.0.0.2:1688`) under
    /// its per-activation-ID registry key, so the renewal task cannot overwrite
    /// a KMS38 (to-2038) lease. `activation_id` is the SPP product ID.
    fn pin_kms38(&self, _activation_id: &str) -> Result<()> {
        Err(Error::UnsupportedPlatform {
            operation: "pin KMS38 lock",
        })
    }
}

/// Construct the SPP backend appropriate for this build.
pub fn backend() -> Box<dyn Spp> {
    #[cfg(all(windows, feature = "winapi"))]
    {
        Box::new(windows_backend::WindowsSpp::new())
    }
    #[cfg(not(all(windows, feature = "winapi")))]
    {
        Box::new(stub::StubSpp)
    }
}
