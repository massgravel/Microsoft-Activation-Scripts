//! KMS38 — extend KMS activation to 2038-01-19.
//!
//! In MAS this is not a separate file: it is a sub-flow of
//! `Online_KMS_Activation.cmd` (with support paths in `Troubleshoot.cmd`). It
//! installs a GVLK, then uses `ClipUp.exe -o` with the special "gatherosstate"
//! ticket so the SPP records an activation whose expiry is pinned to the KMS38
//! epoch instead of the usual 180 days.
//!
//! Port boundary: the epoch/expiry bookkeeping is portable; the actual ticket
//! consumption is the same closed ClipSVC path as [`super::hwid`]. Shares the
//! ClipUp delegation once that slice lands.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

pub struct Kms38;

impl Activator for Kms38 {
    fn method(&self) -> Method {
        Method::Kms38
    }

    fn run(&self, spp: &dyn Spp, _product: Product) -> Result<()> {
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        if spp.windows_build()? < 10240 {
            return Err(Error::Unsupported {
                what: "KMS38 requires Windows 10 or later".into(),
            });
        }
        Err(Error::Unsupported {
            what: "KMS38 ClipUp delegation not yet ported (shares HWID's ClipSVC path)".into(),
        })
    }
}
