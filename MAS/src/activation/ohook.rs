//! `Ohook_Activation_AIO.cmd` — permanent Office activation via SPP hook.
//!
//! **This module cannot be fully reimplemented in safe Rust, by design of the
//! technique.** (Analysis `port_risk`.) Ohook works by placing two
//! reverse-engineered SPP client DLLs (`sppc32.dll` / `sppc64.dll`, acting as
//! `OSPPC.DLL`) that forward to the renamed genuine `sppcs.dll` while overriding
//! the license-status calls so Office always reports "licensed". Those DLLs are
//! opaque binary blobs with **pinned SHA-256 hashes**; a faithful port must ship
//! the identical bytes — Rust cannot "reimplement the SPP internals".
//!
//! What *is* portable and worth porting here:
//!   * the embedded product tables (`OhookEntry`/`MsiEntry`: version, activation
//!     ID GUID, GVLK, license suffix, edition) — pure data, unit-testable;
//!   * build→WMI-class selection (≥9200 ⇒ `SoftwareLicensing*`, else `OSPP*`);
//!   * install-plan assembly: which license `.xrm-ms` files to install and where
//!     to drop the DLL for the detected Office install (MSI vs Click-to-Run vfs).
//!
//! The port therefore = ported planner + verified-hash blob placement + license
//! install through [`Spp::install_license`]. The blobs are assets, not code.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

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
        // The SPP-client blobs are external assets (pinned SHA-256); the planner
        // + license install is the portable slice to port next.
        Err(Error::Unsupported {
            what: "Ohook requires the pinned SPP-client DLL assets (see module docs)".into(),
        })
    }
}
