//! Microsoft Activation Scripts, ported to Rust.
//!
//! A faithful, idiomatic re-architecture of the MAS batch scripts. The design
//! isolates every Windows-specific effect behind the [`platform::Spp`] trait so
//! the portable core — the menu, the domain model, the data tables and the
//! activation orchestration — compiles and unit-tests on any OS, while the real
//! Software Protection Platform work lives in a `--features winapi` Windows
//! backend.
//!
//! ## Module map (→ original scripts)
//! * [`cli`] — the interactive menu (`MAS_AIO.cmd` front-end).
//! * [`model`] — products, methods, license status (the magic GUIDs/strings).
//! * [`platform`] — the SPP boundary + stub / Windows backends.
//! * [`activation::status`] — `Check_Activation_Status.cmd`.
//! * [`activation::online_kms`] — `Online_KMS_Activation.cmd`.
//! * [`activation::hwid`] / [`activation::kms38`] — `HWID_Activation.cmd`.
//! * [`activation::ohook`] — `Ohook_Activation_AIO.cmd`.
//! * [`activation::tsforge`] — `TSforge_Activation.cmd`.
//! * [`data`] — ported data tables.

pub mod activation;
pub mod cli;
pub mod data;
pub mod error;
pub mod model;
pub mod platform;

pub use error::{Error, Result};

/// Entry point used by the binary: build the backend and run the menu.
pub fn run() -> Result<()> {
    let spp = platform::backend();
    cli::run_interactive(spp.as_ref())
}
