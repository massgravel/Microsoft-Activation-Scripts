//! `TSforge_Activation.cmd` — ticket injection into the SPP trusted store.
//!
//! This is the deepest module. TSforge writes activation tickets **directly into
//! the SPP trusted store** (`data.dat`) rather than going through any activation
//! server. Per analysis `port_risk`, the store is a byte-exact,
//! reverse-engineered format: RSA/AES/HMAC-wrapped with CRC32 per block, a
//! `VariableBag` layout, per-block + whole-file SHA-256, physical-store HMAC-SHA1
//! (salted-SHA1 on Vista), and 4/8-byte alignment padding that **differs across
//! Vista / 7 / 8 / 8.1 / Modern**. Any mismatch corrupts the store.
//!
//! A faithful port is essentially porting **LibTSforge** to Rust:
//!   * pure `libtsforge` crate (no_std-friendly, unit-testable on Linux/CI):
//!     the store (de)serialization, CRC/SHA/HMAC layers, key-blob builders, and
//!     the data tables (SKU→edition, MSI-office rows, retail→volume);
//!   * thin `#[cfg(windows)]` FFI: stop `sppsvc`, swap `data.dat`, restart.
//!
//! This is the highest-value, highest-effort slice and the natural next crate to
//! extract. It is intentionally not stubbed as "works" — that would be a lie.

use crate::activation::Activator;
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

pub struct TSforge;

impl Activator for TSforge {
    fn method(&self) -> Method {
        Method::TSforge
    }

    fn run(&self, spp: &dyn Spp, _product: Product) -> Result<()> {
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        Err(Error::Unsupported {
            what: "TSforge requires the SPP trusted-store codec (LibTSforge port — see module docs)"
                .into(),
        })
    }
}
