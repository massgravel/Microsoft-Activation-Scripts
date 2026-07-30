//! Shared in-memory `Spp` fake for unit tests (no Windows required).

use crate::error::Result;
use crate::model::{LicenseStatus, Product};
use crate::platform::{LicenseInfo, Spp};
use std::cell::{Cell, RefCell};
use std::path::Path;

/// Configurable fake SPP. Records calls and lets each test dictate outcomes.
pub struct FakeSpp {
    pub elevated: bool,
    pub build: u32,
    /// ClipSVC produces tokens.dat after the service restart.
    pub tokens_after_restart: bool,
    /// ClipSVC produces tokens.dat after `clipup -v -o`.
    pub tokens_after_clipup: bool,
    /// Products returned by `installed_products`.
    pub products: Vec<LicenseInfo>,
    /// Internal: whether tokens.dat currently "exists" (flipped by restart/clipup).
    pub tokens: Cell<bool>,
    pub calls: RefCell<Vec<String>>,
}

impl Default for FakeSpp {
    fn default() -> Self {
        FakeSpp {
            elevated: true,
            build: 19045,
            tokens_after_restart: false,
            tokens_after_clipup: false,
            products: Vec::new(),
            tokens: Cell::new(false),
            calls: RefCell::new(Vec::new()),
        }
    }
}

impl FakeSpp {
    fn log(&self, s: impl Into<String>) {
        self.calls.borrow_mut().push(s.into());
    }
    pub fn called(&self, s: &str) -> bool {
        self.calls.borrow().iter().any(|c| c == s)
    }
}

impl Spp for FakeSpp {
    fn installed_products(&self, _p: Product) -> Result<Vec<LicenseInfo>> {
        Ok(self.products.clone())
    }
    fn install_product_key(&self, _p: Product, key: &str) -> Result<()> {
        self.log(format!("install_product_key {key}"));
        Ok(())
    }
    fn uninstall_product_key(&self, _p: Product, _id: &str) -> Result<()> {
        Ok(())
    }
    fn set_kms_host(&self, _p: Product, host: &str, port: u16) -> Result<()> {
        self.log(format!("set_kms_host {host}:{port}"));
        Ok(())
    }
    fn clear_kms_host(&self, _p: Product) -> Result<()> {
        Ok(())
    }
    fn activate(&self, _p: Product, id: &str) -> Result<()> {
        self.log(format!("activate {id}"));
        Ok(())
    }
    fn install_license(&self, _x: &Path) -> Result<()> {
        Ok(())
    }
    fn windows_build(&self) -> Result<u32> {
        Ok(self.build)
    }
    fn windows_edition(&self) -> Result<String> {
        Ok("Professional".into())
    }
    fn is_elevated(&self) -> bool {
        self.elevated
    }
    fn generate_genuine_ticket(&self) -> Result<Vec<u8>> {
        self.log("generate_genuine_ticket");
        Ok(b"<genuineTicket/>".to_vec())
    }
    fn write_genuine_ticket(&self, _xml: &[u8]) -> Result<()> {
        self.log("write_genuine_ticket");
        Ok(())
    }
    fn restart_service(&self, name: &str) -> Result<()> {
        self.log(format!("restart_service {name}"));
        if self.tokens_after_restart {
            self.tokens.set(true);
        }
        Ok(())
    }
    fn run_clipup(&self, args: &[&str]) -> Result<()> {
        self.log(format!("run_clipup {}", args.join(" ")));
        if self.tokens_after_clipup {
            self.tokens.set(true);
        }
        Ok(())
    }
    fn clip_tokens_present(&self) -> bool {
        self.tokens.get()
    }
}

/// Build a `LicenseInfo` for tests.
pub fn license(name: &str, key: Option<&str>, status: LicenseStatus) -> LicenseInfo {
    LicenseInfo {
        activation_id: format!("AID-{name}"),
        name: name.into(),
        description: String::new(),
        partial_product_key: key.map(str::to_string),
        status,
        license_family: None,
        grace_minutes: None,
    }
}
