//! Shared trusted-store error type.
//!
//! The concrete container formats live in [`crate::physical_store`] (the
//! Vista/Win7/Modern block dialects) and [`crate::variable_bag`] (the CRC
//! blocks). Both report failures through [`StoreError`].

/// Error decoding a store structure.
#[derive(Debug, PartialEq, Eq)]
pub enum StoreError {
    /// A block's stored CRC did not match its computed CRC (corruption/tamper).
    CrcMismatch { expected: u32, actual: u32 },
    /// Ran off the end of the buffer while decoding.
    Truncated,
}

impl core::fmt::Display for StoreError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            StoreError::CrcMismatch { expected, actual } => write!(
                f,
                "store CRC mismatch (expected 0x{expected:08X}, computed 0x{actual:08X})"
            ),
            StoreError::Truncated => write!(f, "store data truncated"),
        }
    }
}

impl std::error::Error for StoreError {}
