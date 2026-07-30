//! CRC-32 as used by the SPP trusted store (LibTSforge `Utils.CRC32`).
//!
//! This is the **non-reflected CRC-32/BZIP2** variant (poly 0x04C11DB7, init
//! 0xFFFFFFFF, MSB-first, final XOR 0xFFFFFFFF) — *not* the reflected IEEE/zlib
//! CRC (0xEDB88320). Getting this wrong silently corrupts every `CRCBlock` in a
//! `VariableBag`, so it is pinned with the canonical check vector 0xFC891918.

/// CRC-32/BZIP2 of `data`.
pub fn crc32(data: &[u8]) -> u32 {
    let mut crc: u32 = 0xFFFF_FFFF;
    for &b in data {
        // Feed each byte into the HIGH byte and shift left (MSB-first).
        crc ^= (b as u32) << 24;
        for _ in 0..8 {
            crc = if crc & 0x8000_0000 != 0 {
                (crc << 1) ^ 0x04C1_1DB7
            } else {
                crc << 1
            };
        }
    }
    !crc
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bzip2_check_vectors() {
        assert_eq!(crc32(b""), 0x0000_0000);
        assert_eq!(crc32(b"123456789"), 0xFC89_1918); // CRC-32/BZIP2 canonical check
    }

    #[test]
    fn is_not_the_reflected_ieee_crc() {
        // The reflected zlib CRC of "123456789" is 0xCBF43926; ours must differ.
        assert_ne!(crc32(b"123456789"), 0xCBF4_3926);
    }
}
