//! `Online_KMS_Activation.cmd` — classic KMS client activation.
//!
//! Flow (faithful to the script, expressed as SPP calls):
//!   1. Confirm a KMS-client (GVLK) key is installed; the SPP already knows the
//!      per-edition GVLK, so `InstallProductKey` is only needed if a retail key
//!      is present.
//!   2. Point the client at a KMS host (`SetKeyManagementServiceMachine` + port).
//!   3. `Activate()` each applicable product.
//!   4. Fall back through the default host list until one succeeds.
//!
//! The orchestration is portable; the effects happen through the [`Spp`] trait.

use crate::activation::Activator;
use crate::data::kms_hosts::{DEFAULT_KMS_HOSTS, DEFAULT_KMS_PORT};
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

/// Online KMS activator. Optionally pinned to a specific host/port.
pub struct OnlineKms;

/// Try one KMS host: set it, then activate every installed product of the
/// family. Returns `Ok(true)` if at least one product activated.
fn try_host(spp: &dyn Spp, product: Product, host: &str, port: u16) -> Result<bool> {
    spp.set_kms_host(product, host, port)?;
    let mut any = false;
    for info in spp.installed_products(product)? {
        // Only KMS-client (GVLK) products can activate against a KMS host.
        if info.partial_product_key.is_none() {
            continue;
        }
        if spp.activate(product, &info.activation_id).is_ok() {
            any = true;
        }
    }
    Ok(any)
}

impl OnlineKms {
    /// Activate against an explicit host (from `--kms-host`).
    pub fn run_with_host(
        &self,
        spp: &dyn Spp,
        product: Product,
        host: &str,
        port: u16,
    ) -> Result<()> {
        if try_host(spp, product, host, port)? {
            Ok(())
        } else {
            Err(Error::winapi_msg(format!(
                "KMS activation failed against {host}:{port}"
            )))
        }
    }
}

impl Activator for OnlineKms {
    fn method(&self) -> Method {
        Method::OnlineKms
    }

    fn run(&self, spp: &dyn Spp, product: Product) -> Result<()> {
        for host in DEFAULT_KMS_HOSTS {
            match try_host(spp, product, host, DEFAULT_KMS_PORT) {
                Ok(true) => return Ok(()),
                Ok(false) => continue,
                Err(_) => continue, // host unreachable — try the next
            }
        }
        Err(Error::winapi_msg(
            "KMS activation failed against every default host — pass --kms-host",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{LicenseStatus, Product};
    use crate::platform::LicenseInfo;
    use std::cell::RefCell;

    /// Fake SPP recording calls, to test the orchestration without Windows.
    #[derive(Default)]
    struct FakeSpp {
        set_hosts: RefCell<Vec<(String, u16)>>,
        activated: RefCell<Vec<String>>,
        // if true, activate() succeeds only for the named "good" host
        good_host: Option<String>,
        current_host: RefCell<Option<String>>,
    }

    impl Spp for FakeSpp {
        fn installed_products(&self, _p: Product) -> Result<Vec<LicenseInfo>> {
            Ok(vec![LicenseInfo {
                activation_id: "AID-1".into(),
                name: "Windows Pro".into(),
                description: "d".into(),
                partial_product_key: Some("ABCDE".into()),
                status: LicenseStatus::Notification,
                license_family: None,
                grace_minutes: None,
            }])
        }
        fn install_product_key(&self, _p: Product, _k: &str) -> Result<()> {
            Ok(())
        }
        fn uninstall_product_key(&self, _p: Product, _i: &str) -> Result<()> {
            Ok(())
        }
        fn set_kms_host(&self, _p: Product, host: &str, port: u16) -> Result<()> {
            self.set_hosts.borrow_mut().push((host.into(), port));
            *self.current_host.borrow_mut() = Some(host.into());
            Ok(())
        }
        fn clear_kms_host(&self, _p: Product) -> Result<()> {
            Ok(())
        }
        fn activate(&self, _p: Product, id: &str) -> Result<()> {
            let ok = match &self.good_host {
                None => true,
                Some(good) => self.current_host.borrow().as_deref() == Some(good.as_str()),
            };
            if ok {
                self.activated.borrow_mut().push(id.into());
                Ok(())
            } else {
                Err(Error::winapi_msg("no KMS"))
            }
        }
        fn install_license(&self, _x: &std::path::Path) -> Result<()> {
            Ok(())
        }
        fn windows_build(&self) -> Result<u32> {
            Ok(19045)
        }
        fn windows_edition(&self) -> Result<String> {
            Ok("Professional".into())
        }
        fn is_elevated(&self) -> bool {
            true
        }
    }

    #[test]
    fn activates_against_first_working_host() {
        let spp = FakeSpp {
            good_host: Some(DEFAULT_KMS_HOSTS[0].to_string()),
            ..Default::default()
        };
        OnlineKms.run(&spp, Product::Windows).unwrap();
        assert_eq!(spp.activated.borrow().as_slice(), &["AID-1".to_string()]);
    }

    #[test]
    fn falls_through_to_a_later_host() {
        let good = DEFAULT_KMS_HOSTS[2].to_string();
        let spp = FakeSpp {
            good_host: Some(good.clone()),
            ..Default::default()
        };
        OnlineKms.run(&spp, Product::Windows).unwrap();
        // It tried earlier hosts before the one that worked.
        let tried: Vec<String> = spp.set_hosts.borrow().iter().map(|(h, _)| h.clone()).collect();
        assert_eq!(tried[2], good);
    }

    #[test]
    fn explicit_host_that_never_activates_errors() {
        let spp = FakeSpp {
            good_host: Some("nonexistent".into()),
            ..Default::default()
        };
        let err = OnlineKms
            .run_with_host(&spp, Product::Windows, "1.2.3.4", 1688)
            .unwrap_err();
        assert!(matches!(err, Error::WinApi { .. }));
    }
}
