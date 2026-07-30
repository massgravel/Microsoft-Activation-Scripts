//! # libtsforge
//!
//! The SPP trusted-store codec behind TSforge activation, factored out of the
//! `mas` crate so its pure logic is unit-tested on any OS, offline.
//!
//! TSforge writes activation tickets **directly into the Software Protection
//! Platform trusted store** (`data.dat`) rather than contacting an activation
//! server. The deleted `TSforge_Activation.cmd` embeds the complete LibTSforge
//! C# reference implementation, so this is a faithful port of readable source,
//! not a guess. It is layered:
//!
//! * [`crc32`] / [`sha256`] — integrity primitives (in-crate, test-vector
//!   verified; CRC is the non-reflected BZIP2 variant the store actually uses).
//! * [`common`] — `PsVersion` detection, alignment, UTF-16 codec, block-type and
//!   AES-key constants.
//! * [`physical_store`] — the Vista / Win7 / Modern block dialects.
//! * [`variable_bag`] — the two CRC-block dialects (distinct CRC inputs).
//! * [`store`] — the shared error type.
//! * [`tables`] — verbatim product data tables.
//! * [`crypto`] — trait seam for the RSA/AES/HMAC layers a *signed* ticket needs
//!   (behind a future `crypto` feature).
//!
//! Porting status: integrity primitives, `PsVersion`/alignment/UTF-16, the
//! flat physical-store dialects and both VariableBag CRC dialects are ported and
//! round-trip tested. Assembling a *complete signed* ticket additionally needs
//! the Modern physical store, `TokenStoreModern`, the RSA CryptoAPI-blob layer
//! ([`crypto`]) and the verbatim KMS/HWID response blobs — see the module docs.

pub mod common;
pub mod constants;
pub mod crc32;
pub mod crypto;
/// RSA/AES/HMAC backend for signed tickets. Feature-gated: pulls in RustCrypto
/// (`rsa`/`aes`/`cbc`/`hmac`/`sha1`) and does not compile offline.
#[cfg(feature = "crypto")]
pub mod crypto_real;
pub mod physical_store;
pub mod product_key;
pub mod sha256;
pub mod store;
pub mod tables;
pub mod variable_bag;

pub use common::PsVersion;
pub use crc32::crc32;
pub use sha256::sha256;
