//! `VariableBag` CRC blocks (LibTSforge `VariableBag.cs`).
//!
//! A bag is a sequence of key/value entries, each guarded by a CRC-32/BZIP2.
//! There are two dialects with **different byte layouts and different CRC
//! inputs**:
//!
//! * Vista: `DataType, 0, KeyLen, ValueLen, crc, Key, Value` — `crc = CRC32(Value)`.
//! * Modern: `crc, DataType, KeyLen, ValueLen, Key, Align(8), Value, Align(8)` —
//!   but `crc` is computed over a *separate, unaligned* temp buffer
//!   `0i32 ++ DataType ++ KeyLen ++ ValueLen ++ Key ++ Value`. The serialized
//!   bytes are 8-aligned; the CRC input is not. Conflating the two corrupts the
//!   block, so both are tested.

use crate::common::align_pad;
use crate::crc32::crc32;
use crate::store::StoreError;

/// One bag entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CrcBlock {
    /// `CRCBlockType` (Uint=1, String=2, Binary=4).
    pub data_type: u32,
    pub key: Vec<u8>,
    pub value: Vec<u8>,
}

fn u32le(v: u32) -> [u8; 4] {
    v.to_le_bytes()
}

impl CrcBlock {
    /// CRC input for the Modern dialect: `0i32 ++ DataType ++ KeyLen ++
    /// ValueLen ++ Key ++ Value`, with no alignment padding.
    fn modern_crc(&self) -> u32 {
        let mut tmp = Vec::new();
        tmp.extend_from_slice(&u32le(0));
        tmp.extend_from_slice(&u32le(self.data_type));
        tmp.extend_from_slice(&u32le(self.key.len() as u32));
        tmp.extend_from_slice(&u32le(self.value.len() as u32));
        tmp.extend_from_slice(&self.key);
        tmp.extend_from_slice(&self.value);
        crc32(&tmp)
    }

    fn encode_vista(&self, buf: &mut Vec<u8>) {
        buf.extend_from_slice(&u32le(self.data_type));
        buf.extend_from_slice(&u32le(0));
        buf.extend_from_slice(&u32le(self.key.len() as u32));
        buf.extend_from_slice(&u32le(self.value.len() as u32));
        buf.extend_from_slice(&u32le(crc32(&self.value)));
        buf.extend_from_slice(&self.key);
        buf.extend_from_slice(&self.value);
    }

    fn encode_modern(&self, buf: &mut Vec<u8>) {
        buf.extend_from_slice(&u32le(self.modern_crc()));
        buf.extend_from_slice(&u32le(self.data_type));
        buf.extend_from_slice(&u32le(self.key.len() as u32));
        buf.extend_from_slice(&u32le(self.value.len() as u32));
        buf.extend_from_slice(&self.key);
        pad8(buf);
        buf.extend_from_slice(&self.value);
        pad8(buf);
    }
}

fn pad8(buf: &mut Vec<u8>) {
    for _ in 0..align_pad(buf.len(), 8) {
        buf.push(0);
    }
}

fn rd(buf: &[u8], pos: &mut usize, n: usize) -> Result<Vec<u8>, StoreError> {
    if *pos + n > buf.len() {
        return Err(StoreError::Truncated);
    }
    let v = buf[*pos..*pos + n].to_vec();
    *pos += n;
    Ok(v)
}

fn rd_u32(buf: &[u8], pos: &mut usize) -> Result<u32, StoreError> {
    Ok(u32::from_le_bytes(rd(buf, pos, 4)?.try_into().unwrap()))
}

/// Serialize a Vista `VariableBag`.
pub fn encode_bag_vista(blocks: &[CrcBlock]) -> Vec<u8> {
    let mut buf = Vec::new();
    for b in blocks {
        b.encode_vista(&mut buf);
    }
    buf
}

/// Serialize a Modern `VariableBag`.
pub fn encode_bag_modern(blocks: &[CrcBlock]) -> Vec<u8> {
    let mut buf = Vec::new();
    for b in blocks {
        b.encode_modern(&mut buf);
    }
    buf
}

/// Parse a Vista `VariableBag`, verifying each block CRC.
pub fn decode_bag_vista(buf: &[u8]) -> Result<Vec<CrcBlock>, StoreError> {
    let mut pos = 0;
    let mut out = Vec::new();
    while pos + 0x10 <= buf.len() {
        let data_type = rd_u32(buf, &mut pos)?;
        let _zero = rd_u32(buf, &mut pos)?;
        let klen = rd_u32(buf, &mut pos)? as usize;
        let vlen = rd_u32(buf, &mut pos)? as usize;
        let crc = rd_u32(buf, &mut pos)?;
        let key = rd(buf, &mut pos, klen)?;
        let value = rd(buf, &mut pos, vlen)?;
        let actual = crc32(&value);
        if crc != actual {
            return Err(StoreError::CrcMismatch { expected: crc, actual });
        }
        out.push(CrcBlock { data_type, key, value });
    }
    Ok(out)
}

/// Parse a Modern `VariableBag`, verifying each block CRC over the unaligned
/// temp layout.
pub fn decode_bag_modern(buf: &[u8]) -> Result<Vec<CrcBlock>, StoreError> {
    let mut pos = 0;
    let mut out = Vec::new();
    while pos + 0x10 <= buf.len() {
        let crc = rd_u32(buf, &mut pos)?;
        let data_type = rd_u32(buf, &mut pos)?;
        let klen = rd_u32(buf, &mut pos)? as usize;
        let vlen = rd_u32(buf, &mut pos)? as usize;
        let key = rd(buf, &mut pos, klen)?;
        pos += align_pad(pos, 8);
        let value = rd(buf, &mut pos, vlen)?;
        pos += align_pad(pos, 8);
        let block = CrcBlock { data_type, key, value };
        let actual = block.modern_crc();
        if crc != actual {
            return Err(StoreError::CrcMismatch { expected: crc, actual });
        }
        out.push(block);
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn blocks() -> Vec<CrcBlock> {
        vec![
            CrcBlock { data_type: 2, key: b"ProductKey".to_vec(), value: b"XXXXX-YYYYY".to_vec() },
            CrcBlock { data_type: 4, key: b"Pid".to_vec(), value: vec![9, 8, 7, 6, 5] },
        ]
    }

    #[test]
    fn vista_bag_round_trips() {
        let bytes = encode_bag_vista(&blocks());
        assert_eq!(decode_bag_vista(&bytes).unwrap(), blocks());
    }

    #[test]
    fn modern_bag_round_trips_with_alignment() {
        let bytes = encode_bag_modern(&blocks());
        // Every serialized block ends 8-aligned.
        assert_eq!(bytes.len() % 8, 0);
        assert_eq!(decode_bag_modern(&bytes).unwrap(), blocks());
    }

    #[test]
    fn modern_crc_differs_from_vista_crc() {
        // The Modern CRC covers header+key+value; Vista CRC covers value only.
        let b = &blocks()[0];
        assert_ne!(b.modern_crc(), crc32(&b.value));
    }

    #[test]
    fn tampered_value_is_rejected() {
        let mut bytes = encode_bag_vista(&blocks());
        let n = bytes.len();
        bytes[n - 1] ^= 0xFF; // corrupt last value byte
        assert!(matches!(
            decode_bag_vista(&bytes),
            Err(StoreError::CrcMismatch { .. })
        ));
    }
}
