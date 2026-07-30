//! Crypto boundary for the trusted store.
//!
//! CRC-32 and SHA-256 are implemented in-crate ([`crate::crc32`],
//! [`crate::sha256`]). The remaining layers a *signed* ticket needs — HMAC-SHA1
//! for the physical store, RSA to sign the key blob, AES-CBC for the encrypted
//! sections — should be provided by vetted RustCrypto crates behind a future
//! `crypto` feature (`hmac`+`sha1`, `rsa`, `aes`+`cbc`), not hand-rolled.
//!
//! This module defines the trait the store assembler calls, so the rest of the
//! codec is written against a stable interface today and the crate stays
//! dependency-free until the feature is switched on.

/// The asymmetric/keyed operations the store assembler needs but this crate
/// does not yet implement. A `crypto`-feature backend will provide these.
pub trait TicketCrypto {
    /// HMAC-SHA1 over `data` with `key` (physical-store integrity).
    fn hmac_sha1(&self, key: &[u8], data: &[u8]) -> [u8; 20];
    /// AES-128-CBC decrypt (SPP encrypted sections).
    fn aes_cbc_decrypt(&self, key: &[u8], iv: &[u8], data: &[u8]) -> Vec<u8>;
    /// RSA sign `digest` with the embedded private key (key-blob signature).
    fn rsa_sign(&self, digest: &[u8]) -> Vec<u8>;
}
