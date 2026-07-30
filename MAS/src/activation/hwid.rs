//! `HWID_Activation.cmd` — permanent digital-license activation (Windows 10/11).
//!
//! Faithful flow (from the script):
//!   1. Verify build ≥ 10240 and that the edition has a digital-license path.
//!   2. Ensure the correct GVLK/edition key + license files exist under
//!      `%SysPath%\spp\tokens\skus\<edition>\*GVLK*.xrm-ms`.
//!   3. Generate `GenuineTicket.xml` from the machine's hardware hash + region
//!      (`HKCU\Control Panel\International\Geo`) and drop it in
//!      `%ProgramData%\Microsoft\Windows\ClipSVC\GenuineTicket`.
//!   4. Hand the ticket to Windows two ways for reliability: restart ClipSVC,
//!      and run `ClipUp.exe -v -o`. The service consumes the ticket and writes
//!      `tokens.dat`.
//!
//! Port boundary (from analysis `port_risk`): steps 1–3 are pure/portable —
//! the HWID key tables (`HwidEntry`/`FallbackEntry`) and ticket XML assembly are
//! unit-testable data. Step 4 is **closed SPP behavior**: the Rust port cannot
//! reimplement ClipSVC consuming the ticket; like the script, it generates a
//! valid ticket and delegates to `clipup.exe` + the service.
//!
//! Status: portable ticket/table logic is the next slice to port; the ClipUp
//! delegation is a thin [`Spp`]-adjacent shell-out. Not yet wired — running it
//! reports [`Error::Unsupported`] rather than pretending to activate.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

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
        // ponytail: portable ticket-generation + ClipUp delegation not yet
        // ported — this is the next slice, spec captured in the module docs.
        Err(Error::Unsupported {
            what: "HWID ticket generation + ClipUp delegation not yet ported (see module docs)"
                .into(),
        })
    }
}
