//! `HWID_Activation.cmd` — permanent digital-license activation (Windows 10/11).
//!
//! Faithful flow (from the script):
//!   1. Require elevation and build ≥ 10240.
//!   2. Optionally switch the Windows region to the USA (GeoId 244) — the store
//!      license is unavailable in many countries — unless the machine is already
//!      in one of the top countries. Restored afterwards.
//!   3. Build `GenuineTicket.xml` **in-process** — a fixed `<genuineAuthorization>`
//!      XML whose `<properties>` embed a per-edition `Pfn`
//!      (`Microsoft.Windows.{SKU}.{KeyPart}_8wekyb3d8bbwe`) selected from a
//!      34-row key/edition table (+5-row fallback), signed with an embedded
//!      1024-bit RSA key (`clientLockboxKey`) over SHA-256. No `gatherosstate`.
//!   4. Apply it two ways for reliability: restart `ClipSVC`, and if that didn't
//!      produce `tokens.dat`, run `clipup -v -o`. Success = `tokens.dat` exists.
//!
//! Generation is therefore *portable* — given (SKU, KeyPart) a pure builder
//! produces a byte-identical ticket — but it needs the RSA CryptoAPI-blob layer
//! deferred to [`libtsforge::crypto`] (the same blob format TSforge uses), so
//! for now [`Spp::generate_genuine_ticket`] carries it behind the boundary. The
//! region decision and the two-method apply *orchestration* are portable and
//! unit-tested here.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

/// GeoId for the United States (`Set-WinHomeLocation -GeoId 244`).
pub const USA_GEO_ID: u32 = 244;

/// Countries the script does *not* switch away from (store license available).
/// Two-letter geo `Name` values, verbatim from the script.
pub const TOP_COUNTRIES: &[&str] = &[
    "US", "CN", "IN", "BR", "DE", "JP", "GB", "FR", "MX", "ID", "IT", "PK", "TR", "KR", "CA",
    "ES", "AU", "NG", "VN", "PL", "PH", "NL", "EG", "AR", "TH", "CO", "SA", "TW", "MY", "CL",
];

/// Should the region be temporarily switched to the USA before activating?
pub fn should_change_region(current_geo_name: &str) -> bool {
    !TOP_COUNTRIES
        .iter()
        .any(|c| c.eq_ignore_ascii_case(current_geo_name.trim()))
}

/// Generate the ticket, drop it, and apply it via the two ClipSVC methods.
/// Returns `Ok(())` once `tokens.dat` is present.
fn apply_ticket(spp: &dyn Spp) -> Result<()> {
    let ticket = spp.generate_genuine_ticket()?;
    spp.write_genuine_ticket(&ticket)?;

    // Method 1: service restart.
    spp.restart_service("ClipSVC")?;
    if spp.clip_tokens_present() {
        return Ok(());
    }

    // Method 2: clipup -v -o.
    spp.run_clipup(&["-v", "-o"])?;
    if spp.clip_tokens_present() {
        return Ok(());
    }

    Err(Error::winapi_msg(
        "HWID: ClipSVC did not produce tokens.dat after service restart and clipup -v -o",
    ))
}

pub struct Hwid;

impl Activator for Hwid {
    fn method(&self) -> Method {
        Method::Hwid
    }

    fn run(&self, spp: &dyn Spp, _product: Product) -> Result<()> {
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        if spp.windows_build()? < 10240 {
            return Err(Error::Unsupported {
                what: "HWID activation requires Windows 10 or later".into(),
            });
        }
        apply_ticket(spp)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::platform::test_util::FakeSpp;

    #[test]
    fn region_switches_only_outside_top_countries() {
        assert!(!should_change_region("US"));
        assert!(!should_change_region("gb")); // case-insensitive
        assert!(should_change_region("NZ"));
        assert!(should_change_region("SE"));
    }

    #[test]
    fn succeeds_on_service_restart() {
        let spp = FakeSpp {
            tokens_after_restart: true,
            ..Default::default()
        };
        Hwid.run(&spp, Product::Windows).unwrap();
        assert!(spp.called("restart_service ClipSVC"));
        // Never needed the clipup fallback.
        assert!(!spp.called("run_clipup -v -o"));
    }

    #[test]
    fn falls_back_to_clipup() {
        let spp = FakeSpp {
            tokens_after_restart: false,
            tokens_after_clipup: true,
            ..Default::default()
        };
        Hwid.run(&spp, Product::Windows).unwrap();
        assert!(spp.called("run_clipup -v -o"));
    }

    #[test]
    fn errors_when_no_tokens_ever_appear() {
        let spp = FakeSpp::default(); // tokens never appear
        let err = Hwid.run(&spp, Product::Windows).unwrap_err();
        assert!(matches!(err, Error::WinApi { .. }));
    }

    #[test]
    fn requires_elevation_and_win10() {
        let not_elevated = FakeSpp {
            elevated: false,
            ..Default::default()
        };
        assert!(matches!(
            Hwid.run(&not_elevated, Product::Windows),
            Err(Error::NotElevated)
        ));

        let too_old = FakeSpp {
            build: 7601,
            tokens_after_restart: true,
            ..Default::default()
        };
        assert!(matches!(
            Hwid.run(&too_old, Product::Windows),
            Err(Error::Unsupported { .. })
        ));
    }
}
