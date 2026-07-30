//! `TSforge_Activation.cmd` — ticket injection into the SPP trusted store.
//!
//! TSforge writes activation tickets **directly into the SPP trusted store**
//! (`data.dat`) rather than contacting an activation server. The deleted script
//! embeds the full LibTSforge C# reference implementation, which is being ported
//! into the [`libtsforge`] crate:
//!
//! * store dialect selection ([`libtsforge::PsVersion`]),
//! * integrity primitives (CRC-32/BZIP2, SHA-256),
//! * the physical-store and VariableBag block formats.
//!
//! The remaining pieces for a *working* ticket are the Modern physical store,
//! `TokenStoreModern`, the RSA CryptoAPI-blob crypto (behind a `crypto`
//! feature) and the verbatim KMS/HWID response blobs. Until those land this
//! returns a typed error rather than pretending to activate.
//!
//! Pipeline (from the script, per activation ID): detect store version →
//! `InstallGenPKey` → branch `ZeroCID` / `StaticCID` / `KMS4k` → stop `sppsvc`,
//! rewrite the store, restart → verify via WMI.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;
use libtsforge::PsVersion;

pub struct TSforge;

impl Activator for TSforge {
    fn method(&self) -> Method {
        Method::TSforge
    }

    fn run(&self, spp: &dyn Spp, _product: Product) -> Result<()> {
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        // Select the trusted-store dialect for this OS (ported logic).
        let build = spp.windows_build()?;
        let version = PsVersion::detect(build).ok_or_else(|| Error::Unsupported {
            what: format!("TSforge does not support Windows build {build}"),
        })?;

        Err(Error::Unsupported {
            what: format!(
                "TSforge store-write for {version:?} not yet ported \
                 (needs TokenStoreModern + RSA crypto backend — see libtsforge)"
            ),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::platform::test_util::FakeSpp;

    #[test]
    fn requires_elevation() {
        let spp = FakeSpp {
            elevated: false,
            ..Default::default()
        };
        assert!(matches!(
            TSforge.run(&spp, Product::Windows),
            Err(Error::NotElevated)
        ));
    }

    #[test]
    fn selects_store_version_then_reports_unported() {
        // Elevated + a real build → gets past detection to the honest "unported".
        let spp = FakeSpp {
            build: 19045,
            ..Default::default()
        };
        let err = TSforge.run(&spp, Product::Windows).unwrap_err();
        assert!(matches!(err, Error::Unsupported { .. }));
    }

    #[test]
    fn rejects_unsupported_build() {
        let spp = FakeSpp {
            build: 3000, // pre-Vista, no store dialect
            ..Default::default()
        };
        let err = TSforge.run(&spp, Product::Windows).unwrap_err();
        assert!(matches!(err, Error::Unsupported { .. }));
    }
}
