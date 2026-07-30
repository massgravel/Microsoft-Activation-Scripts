//! Shared constants and helpers, ported from LibTSforge `Common.cs` / `Utils.cs`.

/// SPP store generation. Five enum members, but auto-detect never returns
/// `WinBlue` (8.1 → build 9600 → `WinModern`); it exists only for the
/// encryption-version table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PsVersion {
    Vista,
    Win7,
    Win8,
    WinBlue,
    WinModern,
}

impl PsVersion {
    /// `LibTSforge.Utils.DetectVersion()` — keyed off the OS build number.
    /// Returns `None` (C# throws `NotSupportedException`) for unsupported builds.
    pub fn detect(build: u32) -> Option<PsVersion> {
        Some(match build {
            6000..=6003 => PsVersion::Vista,
            7600..=7602 => PsVersion::Win7,
            9200 => PsVersion::Win8,
            b if b >= 9600 => PsVersion::WinModern,
            _ => return None,
        })
    }

    /// The 4-byte version int written at the head of the encrypted physical
    /// store (`PhysStoreCrypto.EncryptPhysicalStore` versionTable).
    pub const fn envelope_version(self) -> u32 {
        match self {
            PsVersion::Vista => 2,
            PsVersion::Win7 => 5,
            PsVersion::Win8 => 1,
            PsVersion::WinBlue => 2,
            PsVersion::WinModern => 3,
        }
    }

    /// Which of the three physical-store dialects this version serializes as.
    pub const fn store_dialect(self) -> StoreDialect {
        match self {
            PsVersion::Vista => StoreDialect::Vista,
            PsVersion::Win7 => StoreDialect::Win7,
            // Win8 / WinBlue / WinModern all use the Modern physical store.
            _ => StoreDialect::Modern,
        }
    }
}

/// The three physical on-disk block dialects (five PS versions collapse to 3).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StoreDialect {
    Vista,
    Win7,
    Modern,
}

/// Physical-store block kind (`BlockType` in Common.cs).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum BlockType {
    None = 0,
    Named = 1,
    Attribute = 2,
    Timer = 3,
}

impl BlockType {
    pub const fn from_u32(v: u32) -> Option<BlockType> {
        Some(match v {
            0 => BlockType::None,
            1 => BlockType::Named,
            2 => BlockType::Attribute,
            3 => BlockType::Timer,
            _ => return None,
        })
    }
}

/// `VariableBag` value type (`CRCBlockType` — bit flags).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum CrcBlockType {
    Uint = 1,
    String = 2,
    Binary = 4,
}

/// The hardcoded AES-128 key for the physical-store envelope
/// (`PhysStoreCrypto`): ASCII `"massgrave.dev :3"`, exactly 16 bytes.
pub const AES_KEY: &[u8; 16] = b"massgrave.dev :3";

/// `BinaryReaderExt.Align(to)` padding: `pad = (-pos) & (to-1)` for power-of-two
/// `to`. Returns the number of padding bytes needed at `pos`.
pub fn align_pad(pos: usize, to: usize) -> usize {
    debug_assert!(to.is_power_of_two());
    pos.wrapping_neg() & (to - 1)
}

/// Encode a string as UTF-16LE with a trailing NUL (`Utils.EncodeString`).
pub fn encode_utf16(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity((s.len() + 1) * 2);
    for u in s.encode_utf16() {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out.extend_from_slice(&[0, 0]); // NUL terminator
    out
}

/// Decode UTF-16LE bytes, dropping a single trailing NUL if present.
pub fn decode_utf16(bytes: &[u8]) -> String {
    let mut units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    if units.last() == Some(&0) {
        units.pop();
    }
    String::from_utf16_lossy(&units)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detect_version_boundaries() {
        assert_eq!(PsVersion::detect(6000), Some(PsVersion::Vista));
        assert_eq!(PsVersion::detect(7601), Some(PsVersion::Win7));
        assert_eq!(PsVersion::detect(9200), Some(PsVersion::Win8));
        assert_eq!(PsVersion::detect(9600), Some(PsVersion::WinModern)); // 8.1 → Modern
        assert_eq!(PsVersion::detect(19045), Some(PsVersion::WinModern));
        assert_eq!(PsVersion::detect(3000), None);
        assert_eq!(PsVersion::detect(9199), None); // gap between Win7 and Win8
    }

    #[test]
    fn envelope_versions_match_the_table() {
        assert_eq!(PsVersion::Vista.envelope_version(), 2);
        assert_eq!(PsVersion::Win7.envelope_version(), 5);
        assert_eq!(PsVersion::Win8.envelope_version(), 1);
        assert_eq!(PsVersion::WinModern.envelope_version(), 3);
    }

    #[test]
    fn dialect_collapse() {
        assert_eq!(PsVersion::Win8.store_dialect(), StoreDialect::Modern);
        assert_eq!(PsVersion::WinBlue.store_dialect(), StoreDialect::Modern);
        assert_eq!(PsVersion::Vista.store_dialect(), StoreDialect::Vista);
    }

    #[test]
    fn align_matches_c_sharp_formula() {
        assert_eq!(align_pad(0, 4), 0);
        assert_eq!(align_pad(1, 4), 3);
        assert_eq!(align_pad(5, 4), 3);
        assert_eq!(align_pad(8, 8), 0);
        assert_eq!(align_pad(9, 8), 7);
    }

    #[test]
    fn utf16_round_trips_with_nul() {
        let enc = encode_utf16("SPPSVC");
        assert_eq!(&enc[enc.len() - 2..], &[0, 0]); // trailing NUL
        assert_eq!(decode_utf16(&enc), "SPPSVC");
    }

    #[test]
    fn aes_key_is_16_bytes() {
        assert_eq!(AES_KEY.len(), 16);
    }
}
