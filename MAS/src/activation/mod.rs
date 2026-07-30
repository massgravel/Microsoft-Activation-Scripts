//! Activation methods.
//!
//! Each activator is a small type implementing [`Activator`]. The menu picks one
//! and calls [`Activator::run`] with the SPP backend and target product. Read-only
//! status reporting lives in [`status`].

use crate::error::Result;
use crate::model::{Method, Product};
use crate::platform::Spp;

pub mod status;

pub mod online_kms;
// Deep SPP-internals activators — scaffolded against the analyzed script specs.
pub mod hwid;
pub mod kms38;
pub mod ohook;
pub mod tsforge;

/// A runnable activation method.
pub trait Activator {
    fn method(&self) -> Method;

    /// Perform the activation against `product`. Implementations must be
    /// idempotent-safe: re-running should not corrupt state (the scripts
    /// re-check status before and after).
    fn run(&self, spp: &dyn Spp, product: Product) -> Result<()>;
}

/// Build the activator for a method.
pub fn activator_for(method: Method) -> Box<dyn Activator> {
    match method {
        Method::OnlineKms => Box::new(online_kms::OnlineKms),
        Method::Hwid => Box::new(hwid::Hwid),
        Method::Kms38 => Box::new(kms38::Kms38),
        Method::Ohook => Box::new(ohook::Ohook),
        Method::TSforge => Box::new(tsforge::TSforge),
    }
}
