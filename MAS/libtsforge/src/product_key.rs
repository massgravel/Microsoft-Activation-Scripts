//! Product-key packing/unpacking and PID helpers, ported from LibTSforge
//! `ProductKey.cs`.
//!
//! A PKEY2009 product key is a 114-bit value packed into two `u64`s
//! (`klow`/`khigh`) carrying the group, serial, security value, an upgrade
//! flag, and a 10-bit CRC-32 checksum, then rendered as the familiar
//! 25-character base-24 string (charset `BCDFGHJKMPQRTVWXY2346789`) with an
//! `N` marker inserted at a position that itself encodes the top digit.
//!
//! This module is pure (std-only): everything the C# reads from the
//! environment (OS build, LCID, date, the non-deterministic PID randomiser) is
//! passed in as a parameter so the logic round-trips under `cargo test`.
//!
//! Intentionally NOT ported (out of scope / not pure):
//! * `GetPkeyId`, `GetPhoneData`, `GetAlgoUri` — belong to the variable-bag /
//!   crypto layers.
//! * the PKEY2005 `ToString` branch — it draws from .NET's seeded `Random`; the
//!   keys it makes are placeholders, and reproducing that PRNG is out of scope.
//! * the `setup.cfg` MPC override in `GetMPC` — file I/O; only the build table
//!   is ported (that is the branch TSforge actually relies on).

use crate::common::encode_utf16;
use crate::crc32::crc32;

/// Base-24 charset (`ProductKey.ALPHABET`). Note: no `N` — the `N` marker in a
/// rendered key is therefore unambiguous.
const ALPHABET: &[u8; 24] = b"BCDFGHJKMPQRTVWXY2346789";

/// `PKeyAlgorithm`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PKeyAlgorithm {
    Pkey2005,
    Pkey2009,
}

/// A decoded product key plus the config-derived metadata the PID builders need.
#[derive(Debug, Clone)]
pub struct ProductKey {
    pub group: u32,
    pub serial: u32,
    pub security: u64,
    pub upgrade: bool,
    pub algorithm: PKeyAlgorithm,
    pub eula_type: String,
    pub part_number: String,
    pub edition: String,
    pub channel: String,
    /// `ActivationId.ToString()` — the config GUID, formatted `d`.
    pub activation_id: String,
    klow: u64,
    khigh: u64,
}

/// `GetMPC` build table (the file-override branch is not ported).
pub fn get_mpc(build: u32) -> &'static str {
    if build >= 10240 {
        "03612"
    } else if build >= 9600 {
        "06401"
    } else if build >= 9200 {
        "05426"
    } else {
        "55041"
    }
}

/// Pack `(group, serial, security, upgrade)` into `(klow, khigh)`, writing the
/// 10-bit CRC-32 checksum into `khigh` bits 39..=48 exactly as the C# ctor does.
pub fn pack_key(group: u32, serial: u32, security: u64, upgrade: bool) -> (u64, u64) {
    let klow = ((security & 0x3fff) << 50)
        | (((serial as u64) & 0x3fff_ffff) << 20)
        | ((group as u64) & 0xf_ffff);
    let mut khigh = ((upgrade as u64) << 49) | ((security >> 14) & 0x7f_ffff_ffff);

    // Checksum is computed over the key bytes *before* it is folded in.
    let mut kb = [0u8; 16];
    kb[..8].copy_from_slice(&klow.to_le_bytes());
    kb[8..].copy_from_slice(&khigh.to_le_bytes());
    let checksum = (crc32(&kb) & 0x3ff) as u64;

    khigh |= checksum << 39;
    (klow, khigh)
}

/// Inverse of [`pack_key`]. Returns `(group, serial, security, upgrade, checksum)`.
pub fn unpack_key(klow: u64, khigh: u64) -> (u32, u32, u64, bool, u32) {
    let group = (klow & 0xf_ffff) as u32;
    let serial = ((klow >> 20) & 0x3fff_ffff) as u32;
    let sec_low = (klow >> 50) & 0x3fff;
    let sec_high = khigh & 0x7f_ffff_ffff;
    let security = (sec_high << 14) | sec_low;
    let upgrade = (khigh >> 49) & 1 == 1;
    let checksum = ((khigh >> 39) & 0x3ff) as u32;
    (group, serial, security, upgrade, checksum)
}

/// Little-endian 15-byte `acc *= 24; acc += d` (top byte stays 0).
fn mul_add_24(acc: &mut [u8; 16], d: u32) {
    let mut carry = d;
    for k in 0..15 {
        let v = acc[k] as u32 * 24 + carry;
        acc[k] = (v & 0xff) as u8;
        carry = v >> 8;
    }
}

/// Base-24 encode a 16-byte key into the dashed 25-char string (PKEY2009 form).
pub fn base24_encode(key: &[u8; 16]) -> String {
    let mut b = *key;
    let mut chars: Vec<u8> = Vec::with_capacity(25);
    let mut last = 0usize;

    for _ in 0..25 {
        let mut current = 0u32;
        for j in (0..15).rev() {
            current = current * 0x100 + b[j] as u32;
            b[j] = (current / 24) as u8;
            current %= 24;
        }
        last = current as usize;
        chars.insert(0, ALPHABET[current as usize]);
    }

    // Drop char 0 (its value is `last`) and reinsert it as the position of `N`.
    let mut out: Vec<u8> = Vec::with_capacity(25);
    out.extend_from_slice(&chars[1..1 + last]);
    out.push(b'N');
    out.extend_from_slice(&chars[last + 1..]);

    out.chunks(5)
        .map(|c| c.iter().map(|&x| x as char).collect::<String>())
        .collect::<Vec<_>>()
        .join("-")
}

/// Inverse of [`base24_encode`]: dashed key string -> 16-byte key.
pub fn base24_decode(s: &str) -> [u8; 16] {
    let clean: Vec<u8> = s.bytes().filter(|&x| x != b'-').collect();
    let last = clean
        .iter()
        .position(|&x| x == b'N')
        .expect("product key has no 'N' marker");

    let mut acc = [0u8; 16];
    // digits, most-significant first: [last] then every non-'N' char in order.
    mul_add_24(&mut acc, last as u32);
    for (i, &c) in clean.iter().enumerate() {
        if i == last {
            continue;
        }
        let d = ALPHABET
            .iter()
            .position(|&a| a == c)
            .expect("invalid base-24 character") as u32;
        mul_add_24(&mut acc, d);
    }
    acc
}

impl ProductKey {
    /// Mirror of the C# constructor (minus the config/range plumbing): pack the
    /// key and stash the metadata the PID builders reference.
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        group: u32,
        serial: u32,
        security: u64,
        upgrade: bool,
        algorithm: PKeyAlgorithm,
        eula_type: impl Into<String>,
        part_number: impl Into<String>,
        edition: impl Into<String>,
        channel: impl Into<String>,
        activation_id: impl Into<String>,
    ) -> Self {
        let (klow, khigh) = pack_key(group, serial, security, upgrade);
        ProductKey {
            group,
            serial,
            security,
            upgrade,
            algorithm,
            eula_type: eula_type.into(),
            part_number: part_number.into(),
            edition: edition.into(),
            channel: channel.into(),
            activation_id: activation_id.into(),
            klow,
            khigh,
        }
    }

    /// `KeyBytes` = `klow` ‖ `khigh`, both little-endian (16 bytes).
    pub fn key_bytes(&self) -> [u8; 16] {
        let mut kb = [0u8; 16];
        kb[..8].copy_from_slice(&self.klow.to_le_bytes());
        kb[8..].copy_from_slice(&self.khigh.to_le_bytes());
        kb
    }

    /// The 10-bit checksum packed into `khigh`.
    pub fn checksum(&self) -> u32 {
        ((self.khigh >> 39) & 0x3ff) as u32
    }

    /// Rendered PKEY2009 product-key string.
    pub fn to_key_string(&self) -> String {
        debug_assert_eq!(self.algorithm, PKeyAlgorithm::Pkey2009);
        base24_encode(&self.key_bytes())
    }

    /// `GetPid2` (only PKEY2005 produces a value; PKEY2009 returns `""`).
    /// `rand_1000` supplies the `Random().Next(1000)` term of the non-OEM path.
    pub fn get_pid2(&self, build: u32, rand_1000: u32) -> String {
        if self.algorithm != PKeyAlgorithm::Pkey2005 {
            return String::new();
        }
        let mpc = get_mpc(build);
        let (serial_high, serial_low, last_part): (String, u32, u32) = if self.eula_type == "OEM" {
            (
                "OEM".to_string(),
                (self.group / 2 % 100) * 10000 + self.serial / 100000,
                self.serial % 100000,
            )
        } else {
            (
                format!("{:03}", self.serial / 1000000),
                self.serial % 1000000,
                (self.group / 2 % 100) * 1000 + rand_1000,
            )
        };

        let digit_sum: u32 = serial_low
            .to_string()
            .bytes()
            .map(|b| (b - b'0') as u32)
            .sum();
        let checksum = 7 - (digit_sum % 7);

        format!(
            "{}-{}-{:06}{}-{:05}",
            mpc, serial_high, serial_low, checksum, last_part
        )
    }

    /// `GetPid3`.
    pub fn get_pid3(&self, build: u32, rand_1000: u32) -> Vec<u8> {
        let mut out: Vec<u8> = Vec::new();
        out.extend_from_slice(&0xA4u32.to_le_bytes());
        out.extend_from_slice(&0x3u32.to_le_bytes());
        write_fixed_ascii(&mut out, &self.get_pid2(build, rand_1000), 24);
        out.extend_from_slice(&self.group.to_le_bytes());
        write_fixed_ascii(&mut out, &self.part_number, 16);
        out.extend(std::iter::repeat(0u8).take(0x6C));

        let mut rev = out.clone();
        rev.reverse();
        let mut crc = (!crc32(&rev)).to_le_bytes();
        crc.reverse();
        out.extend_from_slice(&crc);
        out
    }

    /// `GetExtendedPid`.
    pub fn get_extended_pid(&self, build: u32, lcid: u32, day_of_year: u32, year: u32) -> String {
        let mpc = get_mpc(build);
        let serial_high = self.serial / 1000000;
        let serial_low = self.serial % 1000000;
        let license_type = match self.eula_type.as_str() {
            "OEM" => 2,
            "Volume" => 3,
            _ => 0,
        };
        format!(
            "{}-{:05}-{:03}-{:06}-{:02}-{:04}-{:04}.0000-{:03}{:04}",
            mpc, self.group, serial_high, serial_low, license_type, lcid, build, day_of_year, year
        )
    }

    /// `GetPid4`.
    pub fn get_pid4(&self, build: u32, lcid: u32, day_of_year: u32, year: u32) -> Vec<u8> {
        let mut out: Vec<u8> = Vec::new();
        out.extend_from_slice(&0x4F8u32.to_le_bytes());
        out.extend_from_slice(&0x4u32.to_le_bytes());
        write_fixed_utf16(
            &mut out,
            &self.get_extended_pid(build, lcid, day_of_year, year),
            0x80,
        );
        write_fixed_utf16(&mut out, &self.activation_id, 0x80);
        out.extend(std::iter::repeat(0u8).take(0x10));
        write_fixed_utf16(&mut out, &self.edition, 0x208);
        out.extend_from_slice(&(self.upgrade as u64).to_le_bytes());
        out.extend(std::iter::repeat(0u8).take(0x50));
        write_fixed_utf16(&mut out, &self.part_number, 0x80);
        write_fixed_utf16(&mut out, &self.channel, 0x80);
        write_fixed_utf16(&mut out, &self.eula_type, 0x80);
        out
    }
}

/// `WriteFixedString`: ASCII bytes zero-padded to `blen`.
fn write_fixed_ascii(out: &mut Vec<u8>, s: &str, blen: usize) {
    out.extend_from_slice(s.as_bytes());
    out.extend(std::iter::repeat(0u8).take(blen - s.len()));
}

/// `WriteFixedString16`: UTF-16LE + NUL, zero-padded to `blen`.
fn write_fixed_utf16(out: &mut Vec<u8>, s: &str, blen: usize) {
    let enc = encode_utf16(s);
    out.extend_from_slice(&enc);
    out.extend(std::iter::repeat(0u8).take(blen - enc.len()));
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_2009() -> ProductKey {
        ProductKey::new(
            2265,      // group
            123456789, // serial
            0x1_2345,  // security
            true,      // upgrade
            PKeyAlgorithm::Pkey2009,
            "OEM",
            "X19-98765",
            "Professional",
            "Retail",
            "12345678-1234-1234-1234-123456789012",
        )
    }

    #[test]
    fn base24_round_trips() {
        let pk = sample_2009();
        let s = pk.to_key_string();

        // Rendered shape: 5 groups of 5, exactly one 'N', charset-clean.
        assert_eq!(s.len(), 29); // 25 chars + 4 dashes
        assert_eq!(s.matches('N').count(), 1);
        assert!(s
            .bytes()
            .all(|b| b == b'-' || b == b'N' || ALPHABET.contains(&b)));

        assert_eq!(base24_decode(&s), pk.key_bytes());
    }

    #[test]
    fn base24_decode_ignores_grouping() {
        let pk = sample_2009();
        let dashed = pk.to_key_string();
        let undashed: String = dashed.chars().filter(|&c| c != '-').collect();
        assert_eq!(base24_decode(&dashed), base24_decode(&undashed));
    }

    #[test]
    fn pack_unpack_round_trips() {
        let (klow, khigh) = pack_key(2265, 123456789, 0x1_2345, true);
        let (g, s, sec, up, _cs) = unpack_key(klow, khigh);
        assert_eq!((g, s, sec, up), (2265, 123456789, 0x1_2345, true));
    }

    #[test]
    fn checksum_lands_in_khigh_bits_39_to_48() {
        let pk = sample_2009();
        let kb = pk.key_bytes();
        let khigh = u64::from_le_bytes(kb[8..].try_into().unwrap());
        let field = ((khigh >> 39) & 0x3ff) as u32;

        // Recompute over the key with the checksum field zeroed.
        let khigh_no_cs = khigh & !(0x3ffu64 << 39);
        let mut kb2 = kb;
        kb2[8..].copy_from_slice(&khigh_no_cs.to_le_bytes());
        assert_eq!(field, crc32(&kb2) & 0x3ff);
        assert_eq!(field, pk.checksum());
    }

    #[test]
    fn mpc_by_build() {
        assert_eq!(get_mpc(19045), "03612");
        assert_eq!(get_mpc(10240), "03612");
        assert_eq!(get_mpc(10239), "06401");
        assert_eq!(get_mpc(9600), "06401");
        assert_eq!(get_mpc(9200), "05426");
        assert_eq!(get_mpc(7601), "55041");
    }

    #[test]
    fn pid_blobs_have_the_expected_fixed_sizes() {
        let pk = sample_2009();
        // 4 + 4 + 24 + 4 + 16 + 108 + 4(crc)
        assert_eq!(pk.get_pid3(19045, 0).len(), 164);
        // header value 0x4F8 == total length.
        assert_eq!(pk.get_pid4(19045, 1033, 210, 2026).len(), 0x4F8);
    }

    #[test]
    fn pid2_oem_checksum_and_shape() {
        // A PKEY2005 OEM key exercises the deterministic PID2 path.
        let pk = ProductKey::new(
            100,
            250123,
            0,
            false,
            PKeyAlgorithm::Pkey2005,
            "OEM",
            "PN",
            "Core",
            "OEM",
            "id",
        );
        let pid2 = pk.get_pid2(9200, 0);
        // mpc-OEM-{serialLow:06}{chk}-{lastPart:05}
        // serialLow = (100/2 % 100)*10000 + 250123/100000 = 500000 + 2 = 500002
        // digitsum(500002)=7 -> chk = 7 - 0 = 7 ; lastPart = 250123 % 100000 = 50123
        assert_eq!(pid2, "05426-OEM-5000027-50123");
    }
}
