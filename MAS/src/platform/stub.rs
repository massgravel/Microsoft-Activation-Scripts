//! Non-Windows stub backend.
//!
//! Lets the entire crate compile and unit-test off Windows. Read-only queries
//! return empty/neutral data; anything that would mutate activation state fails
//! loudly with [`Error::UnsupportedPlatform`] instead of silently pretending to
//! work.

use crate::error::{Error, Result};
use crate::model::Product;
use crate::platform::{LicenseInfo, Spp};
use std::path::Path;

#[derive(Default)]
pub struct StubSpp;

impl StubSpp {
    fn deny(op: &'static str) -> Error {
        Error::UnsupportedPlatform { operation: op }
    }
}

impl Spp for StubSpp {
    fn installed_products(&self, _product: Product) -> Result<Vec<LicenseInfo>> {
        Ok(Vec::new())
    }

    fn install_product_key(&self, _p: Product, _key: &str) -> Result<()> {
        Err(Self::deny("install product key"))
    }

    fn uninstall_product_key(&self, _p: Product, _id: &str) -> Result<()> {
        Err(Self::deny("uninstall product key"))
    }

    fn set_kms_host(&self, _p: Product, _host: &str, _port: u16) -> Result<()> {
        Err(Self::deny("set KMS host"))
    }

    fn clear_kms_host(&self, _p: Product) -> Result<()> {
        Err(Self::deny("clear KMS host"))
    }

    fn activate(&self, _p: Product, _id: &str) -> Result<()> {
        Err(Self::deny("activate product"))
    }

    fn install_license(&self, _xml_path: &Path) -> Result<()> {
        Err(Self::deny("install license"))
    }

    fn windows_build(&self) -> Result<u32> {
        // No Windows to ask; report 0 so build-gated methods are filtered out.
        Ok(0)
    }

    fn windows_edition(&self) -> Result<String> {
        Ok("Unknown".to_string())
    }

    fn is_elevated(&self) -> bool {
        false
    }
}
