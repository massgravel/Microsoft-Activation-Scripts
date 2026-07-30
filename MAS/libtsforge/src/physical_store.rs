//! Physical-store block dialects (LibTSforge `PhysicalStoreVista/Win7/Modern`).
//!
//! The decrypted physical store is `8` pre-header bytes followed by a list of
//! blocks, each `Align(4)`-padded. Vista and Win7 are flat block lists (ported
//! and round-trip-tested here); the Modern dialect groups blocks by UTF-16 key
//! (scaffolded — see [`StoreDialect::Modern`]).
//!
//! Note on fidelity: without a real `data.dat` these tests prove encode/decode
//! symmetry, not byte-equality with Windows. The exact trailing-slack bound the
//! real reader uses (`pos < len - 0x14`) is documented on [`decode_flat`].

use crate::common::{align_pad, BlockType, StoreDialect};
use crate::store::StoreError;

/// One physical-store record. `key` is empty in the Vista dialect.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PsBlock {
    pub ty: BlockType,
    pub flags: u32,
    pub key: Vec<u8>,
    pub value: Vec<u8>,
    pub data: Vec<u8>,
}

const PREHEADER_LEN: usize = 8;

fn put_u32(buf: &mut Vec<u8>, v: u32) {
    buf.extend_from_slice(&v.to_le_bytes());
}

fn read_u32(buf: &[u8], pos: &mut usize) -> Result<u32, StoreError> {
    if *pos + 4 > buf.len() {
        return Err(StoreError::Truncated);
    }
    let v = u32::from_le_bytes(buf[*pos..*pos + 4].try_into().unwrap());
    *pos += 4;
    Ok(v)
}

fn read_bytes(buf: &[u8], pos: &mut usize, len: usize) -> Result<Vec<u8>, StoreError> {
    if *pos + len > buf.len() {
        return Err(StoreError::Truncated);
    }
    let out = buf[*pos..*pos + len].to_vec();
    *pos += len;
    Ok(out)
}

fn pad4(buf: &mut Vec<u8>) {
    for _ in 0..align_pad(buf.len(), 4) {
        buf.push(0);
    }
}

fn skip_align4(pos: &mut usize) {
    *pos += align_pad(*pos, 4);
}

/// Serialize a flat block list (Vista or Win7 dialect) with the 8-byte
/// pre-header and 4-byte inter-block alignment.
pub fn encode_flat(preheader: &[u8; PREHEADER_LEN], blocks: &[PsBlock], dialect: StoreDialect) -> Vec<u8> {
    let mut buf = Vec::new();
    buf.extend_from_slice(preheader);
    for b in blocks {
        put_u32(&mut buf, b.ty as u32);
        put_u32(&mut buf, b.flags);
        match dialect {
            StoreDialect::Vista => {
                // Type, Flags, Value.Length, Data.Length, Value, Data (no key).
                put_u32(&mut buf, b.value.len() as u32);
                put_u32(&mut buf, b.data.len() as u32);
                buf.extend_from_slice(&b.value);
                buf.extend_from_slice(&b.data);
            }
            StoreDialect::Win7 => {
                // Type, Flags, Key.Length, Value.Length, Data.Length, Key, Value, Data.
                put_u32(&mut buf, b.key.len() as u32);
                put_u32(&mut buf, b.value.len() as u32);
                put_u32(&mut buf, b.data.len() as u32);
                buf.extend_from_slice(&b.key);
                buf.extend_from_slice(&b.value);
                buf.extend_from_slice(&b.data);
            }
            StoreDialect::Modern => unreachable!("Modern uses encode_modern"),
        }
        pad4(&mut buf);
    }
    buf
}

/// Deserialize a flat block list. Stops when fewer than a minimal header
/// remains. The real Windows reader loops while `pos < len - 0x14`; here we
/// stop symmetrically with what [`encode_flat`] wrote.
pub fn decode_flat(buf: &[u8], dialect: StoreDialect) -> Result<([u8; PREHEADER_LEN], Vec<PsBlock>), StoreError> {
    if buf.len() < PREHEADER_LEN {
        return Err(StoreError::Truncated);
    }
    let mut preheader = [0u8; PREHEADER_LEN];
    preheader.copy_from_slice(&buf[..PREHEADER_LEN]);
    let mut pos = PREHEADER_LEN;

    // Smallest header: Vista = 4 u32 (0x10), Win7 = 5 u32 (0x14).
    let min_header = match dialect {
        StoreDialect::Vista => 16,
        StoreDialect::Win7 => 20,
        StoreDialect::Modern => return Err(StoreError::Truncated),
    };

    let mut blocks = Vec::new();
    while pos + min_header <= buf.len() {
        let ty_raw = read_u32(buf, &mut pos)?;
        let ty = BlockType::from_u32(ty_raw).ok_or(StoreError::Truncated)?;
        let flags = read_u32(buf, &mut pos)?;
        let (key, value, data) = match dialect {
            StoreDialect::Vista => {
                let vlen = read_u32(buf, &mut pos)? as usize;
                let dlen = read_u32(buf, &mut pos)? as usize;
                let value = read_bytes(buf, &mut pos, vlen)?;
                let data = read_bytes(buf, &mut pos, dlen)?;
                (Vec::new(), value, data)
            }
            StoreDialect::Win7 => {
                let klen = read_u32(buf, &mut pos)? as usize;
                let vlen = read_u32(buf, &mut pos)? as usize;
                let dlen = read_u32(buf, &mut pos)? as usize;
                let key = read_bytes(buf, &mut pos, klen)?;
                let value = read_bytes(buf, &mut pos, vlen)?;
                let data = read_bytes(buf, &mut pos, dlen)?;
                (key, value, data)
            }
            StoreDialect::Modern => unreachable!(),
        };
        blocks.push(PsBlock { ty, flags, key, value, data });
        skip_align4(&mut pos);
    }
    Ok((preheader, blocks))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_blocks() -> Vec<PsBlock> {
        vec![
            PsBlock {
                ty: BlockType::Named,
                flags: 0x402,
                key: b"appId".to_vec(),
                value: b"pkeyId-value".to_vec(),
                data: vec![1, 2, 3],
            },
            PsBlock {
                ty: BlockType::Timer,
                flags: 0x4,
                key: b"k2".to_vec(),
                value: b"v".to_vec(),
                data: vec![],
            },
        ]
    }

    #[test]
    fn vista_round_trips() {
        let pre = [0xAAu8; 8];
        // Vista carries no key; clear it so equality holds.
        let blocks: Vec<PsBlock> = sample_blocks()
            .into_iter()
            .map(|mut b| {
                b.key = Vec::new();
                b
            })
            .collect();
        let bytes = encode_flat(&pre, &blocks, StoreDialect::Vista);
        assert_eq!(bytes.len() % 4, 0); // 4-byte aligned
        let (got_pre, got) = decode_flat(&bytes, StoreDialect::Vista).unwrap();
        assert_eq!(got_pre, pre);
        assert_eq!(got, blocks);
    }

    #[test]
    fn win7_round_trips_with_keys() {
        let pre = [0u8; 8];
        let blocks = sample_blocks();
        let bytes = encode_flat(&pre, &blocks, StoreDialect::Win7);
        let (_, got) = decode_flat(&bytes, StoreDialect::Win7).unwrap();
        assert_eq!(got, blocks);
    }

    #[test]
    fn truncated_preheader_errs() {
        assert_eq!(decode_flat(b"\x00\x00", StoreDialect::Vista), Err(StoreError::Truncated));
    }
}
