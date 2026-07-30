//! HWID key/edition tables, generated verbatim from `HWID_Activation.cmd`
//! (`:hwiddata` / `:hwidfallback`, git history `f34d025`).
//!
//! The script stores keys with a `%f%` empty-var splitter as light
//! obfuscation; it is stripped here. Rows are keyed by the SPP SKU id, which
//! is matched against the running edition's SKU. SKU ids are **not unique**
//! (e.g. the EnterpriseS branches all share SKU 125); the script disambiguates
//! by OS build/version, so [`entry_for_sku`] returns the first match and
//! callers refine by `version` when needed.

/// One HWID key/edition row.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HwidEntry {
    pub sku: u32,
    pub activation_id: &'static str,
    /// GVLK/generic product key (the `%f%` splitter already stripped).
    pub product_key: &'static str,
    pub key_part: &'static str,
    /// `false` if the script marks this key as not working (`%%E == 1`).
    pub works: bool,
    pub key_type: &'static str,
    pub edition_id: &'static str,
    /// Build/branch tag, or empty (e.g. `RS5`, `Ge`, `Zn`).
    pub version: &'static str,
}

/// Alternate-edition fallback: when the current edition lacks a working HWID
/// key, activate as the mapped alternate edition instead.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct HwidFallback {
    pub cur_sku: u32,
    pub cur_edition: &'static str,
    pub cur_activation_id: &'static str,
    pub alt_activation_id: &'static str,
    pub alt_key: &'static str,
    pub alt_edition: &'static str,
}

/// The 34 HWID key/edition rows, verbatim.
pub const HWID_KEYS: &[HwidEntry] = &[
    HwidEntry { sku: 4, activation_id: "8b351c9c-f398-4515-9900-09df49427262", product_key: "XGVPP-NMH47-7TTHJ-W3FW7-8HV2C", key_part: "X19-99683", works: true, key_type: "OEM:NONSLP", edition_id: "Enterprise", version: "" },
    HwidEntry { sku: 27, activation_id: "c83cef07-6b72-4bbc-a28f-a00386872839", product_key: "3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT", key_part: "X19-98746", works: true, key_type: "Volume:MAK", edition_id: "EnterpriseN", version: "" },
    HwidEntry { sku: 48, activation_id: "4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c", product_key: "VK7JG-NPHTM-C97JM-9MPGT-3V66T", key_part: "X19-98841", works: true, key_type: "Retail", edition_id: "Professional", version: "" },
    HwidEntry { sku: 49, activation_id: "9fbaf5d6-4d83-4422-870d-fdda6e5858aa", product_key: "2B87N-8KFHP-DKV6R-Y2C8J-PKCKT", key_part: "X19-98859", works: true, key_type: "Retail", edition_id: "ProfessionalN", version: "" },
    HwidEntry { sku: 98, activation_id: "f742e4ff-909d-4fe9-aacb-3231d24a0c58", product_key: "4CPRK-NM3K3-X6XXQ-RXX86-WXCHW", key_part: "X19-98877", works: true, key_type: "Retail", edition_id: "CoreN", version: "" },
    HwidEntry { sku: 99, activation_id: "1d1bac85-7365-4fea-949a-96978ec91ae0", product_key: "N2434-X9D7W-8PF6X-8DV9T-8TYMD", key_part: "X19-99652", works: true, key_type: "Retail", edition_id: "CoreCountrySpecific", version: "" },
    HwidEntry { sku: 100, activation_id: "3ae2cc14-ab2d-41f4-972f-5e20142771dc", product_key: "BT79Q-G7N6G-PGBYW-4YWX6-6F4BT", key_part: "X19-99661", works: true, key_type: "Retail", edition_id: "CoreSingleLanguage", version: "" },
    HwidEntry { sku: 101, activation_id: "2b1f36bb-c1cd-4306-bf5c-a0367c2d97d8", product_key: "YTMG3-N6DKC-DKB77-7M9GH-8HVX7", key_part: "X19-98868", works: true, key_type: "Retail", edition_id: "Core", version: "" },
    HwidEntry { sku: 119, activation_id: "2a6137f3-75c0-4f26-8e3e-d83d802865a4", product_key: "XKCNC-J26Q9-KFHD2-FKTHY-KD72Y", key_part: "X19-99606", works: true, key_type: "OEM:NONSLP", edition_id: "PPIPro", version: "" },
    HwidEntry { sku: 121, activation_id: "e558417a-5123-4f6f-91e7-385c1c7ca9d4", product_key: "YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY", key_part: "X19-98886", works: true, key_type: "Retail", edition_id: "Education", version: "" },
    HwidEntry { sku: 122, activation_id: "c5198a66-e435-4432-89cf-ec777c9d0352", product_key: "84NGF-MHBT6-FXBX8-QWJK7-DRR8H", key_part: "X19-98892", works: true, key_type: "Retail", edition_id: "EducationN", version: "" },
    HwidEntry { sku: 125, activation_id: "f6e29426-a256-4316-88bf-cc5b0f95ec0c", product_key: "PJB47-8PN2T-MCGDY-JTY3D-CBCPV", key_part: "X23-50331", works: false, key_type: "Volume:MAK", edition_id: "EnterpriseS", version: "Ge" },
    HwidEntry { sku: 125, activation_id: "cce9d2de-98ee-4ce2-8113-222620c64a27", product_key: "KCNVH-YKWX8-GJJB9-H9FDT-6F7W2", key_part: "X22-66075", works: false, key_type: "Volume:MAK", edition_id: "EnterpriseS", version: "VB" },
    HwidEntry { sku: 125, activation_id: "d06934ee-5448-4fd1-964a-cd077618aa06", product_key: "43TBQ-NH92J-XKTM7-KT3KK-P39PB", key_part: "X21-83233", works: true, key_type: "OEM:NONSLP", edition_id: "EnterpriseS", version: "RS5" },
    HwidEntry { sku: 125, activation_id: "706e0cfd-23f4-43bb-a9af-1a492b9f1302", product_key: "NK96Y-D9CD8-W44CQ-R8YTK-DYJWX", key_part: "X21-05035", works: true, key_type: "OEM:NONSLP", edition_id: "EnterpriseS", version: "RS1" },
    HwidEntry { sku: 125, activation_id: "faa57748-75c8-40a2-b851-71ce92aa8b45", product_key: "FWN7H-PF93Q-4GGP8-M8RF3-MDWWW", key_part: "X19-99617", works: true, key_type: "OEM:NONSLP", edition_id: "EnterpriseS", version: "TH" },
    HwidEntry { sku: 126, activation_id: "3d1022d8-969f-4222-b54b-327f5a5af4c9", product_key: "2DBW3-N2PJG-MVHW3-G7TDK-9HKR4", key_part: "X21-04921", works: true, key_type: "Volume:MAK", edition_id: "EnterpriseSN", version: "RS1" },
    HwidEntry { sku: 126, activation_id: "60c243e1-f90b-4a1b-ba89-387294948fb6", product_key: "NTX6B-BRYC2-K6786-F6MVQ-M7V2X", key_part: "X19-98770", works: true, key_type: "Volume:MAK", edition_id: "EnterpriseSN", version: "TH" },
    HwidEntry { sku: 139, activation_id: "01eb852c-424d-4060-94b8-c10d799d7364", product_key: "3XP6D-CRND4-DRYM2-GM84D-4GG8Y", key_part: "X23-37869", works: false, key_type: "Retail", edition_id: "ProfessionalCountrySpecific", version: "Zn" },
    HwidEntry { sku: 161, activation_id: "eb6d346f-1c60-4643-b960-40ec31596c45", product_key: "DXG7C-N36C4-C4HTG-X4T3X-2YV77", key_part: "X21-43626", works: true, key_type: "Retail", edition_id: "ProfessionalWorkstation", version: "" },
    HwidEntry { sku: 162, activation_id: "89e87510-ba92-45f6-8329-3afa905e3e83", product_key: "WYPNQ-8C467-V2W6J-TX4WX-WT2RQ", key_part: "X21-43644", works: true, key_type: "Retail", edition_id: "ProfessionalWorkstationN", version: "" },
    HwidEntry { sku: 164, activation_id: "62f0c100-9c53-4e02-b886-a3528ddfe7f6", product_key: "8PTT6-RNW4C-6V7J2-C2D3X-MHBPB", key_part: "X21-04955", works: true, key_type: "Retail", edition_id: "ProfessionalEducation", version: "" },
    HwidEntry { sku: 165, activation_id: "13a38698-4a49-4b9e-8e83-98fe51110953", product_key: "GJTYN-HDMQY-FRR76-HVGC7-QPF8P", key_part: "X21-04956", works: true, key_type: "Retail", edition_id: "ProfessionalEducationN", version: "" },
    HwidEntry { sku: 175, activation_id: "df96023b-dcd9-4be2-afa0-c6c871159ebe", product_key: "NJCF7-PW8QT-3324D-688JX-2YV66", key_part: "X21-41295", works: true, key_type: "Retail", edition_id: "ServerRdsh", version: "" },
    HwidEntry { sku: 178, activation_id: "d4ef7282-3d2c-4cf0-9976-8854e64a8d1e", product_key: "V3WVW-N2PV2-CGWC3-34QGF-VMJ2C", key_part: "X21-32983", works: true, key_type: "Retail", edition_id: "Cloud", version: "" },
    HwidEntry { sku: 179, activation_id: "af5c9381-9240-417d-8d35-eb40cd03e484", product_key: "NH9J3-68WK7-6FB93-4K3DF-DJ4F6", key_part: "X21-32987", works: true, key_type: "Retail", edition_id: "CloudN", version: "" },
    HwidEntry { sku: 188, activation_id: "8ab9bdd1-1f67-4997-82d9-8878520837d9", product_key: "XQQYW-NFFMW-XJPBH-K8732-CKFFD", key_part: "X21-99378", works: true, key_type: "OEM:DM", edition_id: "IoTEnterprise", version: "" },
    HwidEntry { sku: 191, activation_id: "ed655016-a9e8-4434-95d9-4345352c2552", product_key: "QPM6N-7J2WJ-P88HH-P3YRH-YY74H", key_part: "X21-99682", works: true, key_type: "OEM:NONSLP", edition_id: "IoTEnterpriseS", version: "VB" },
    HwidEntry { sku: 191, activation_id: "6c4de1b8-24bb-4c17-9a77-7b939414c298", product_key: "CGK42-GYN6Y-VD22B-BX98W-J8JXD", key_part: "X23-12617", works: true, key_type: "OEM:NONSLP", edition_id: "IoTEnterpriseS", version: "Ge" },
    HwidEntry { sku: 202, activation_id: "d4bdc678-0a4b-4a32-a5b3-aaa24c3b0f24", product_key: "K9VKN-3BGWV-Y624W-MCRMQ-BHDCD", key_part: "X22-53884", works: true, key_type: "Retail", edition_id: "CloudEditionN", version: "" },
    HwidEntry { sku: 203, activation_id: "92fb8726-92a8-4ffc-94ce-f82e07444653", product_key: "KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W", key_part: "X22-53847", works: true, key_type: "Retail", edition_id: "CloudEdition", version: "" },
    HwidEntry { sku: 205, activation_id: "5a85300a-bfce-474f-ac07-a30983e3fb90", product_key: "N979K-XWD77-YW3GB-HBGH6-D32MH", key_part: "X23-15042", works: true, key_type: "OEM:DM", edition_id: "IoTEnterpriseSK", version: "" },
    HwidEntry { sku: 206, activation_id: "80083eae-7031-4394-9e88-4901973d56fe", product_key: "P8Q7T-WNK7X-PMFXY-VXHBG-RRK69", key_part: "X23-62084", works: true, key_type: "OEM:DM", edition_id: "IoTEnterpriseK", version: "" },
    HwidEntry { sku: 210, activation_id: "1bc2140b-285b-4351-b99c-26a126104b29", product_key: "TMP2N-KGFHJ-PWM6F-68KCQ-3PJBP", key_part: "X23-60513", works: true, key_type: "Retail", edition_id: "WNC", version: "" },
];

/// The 5 alternate-edition fallback rows, verbatim.
pub const HWID_FALLBACK: &[HwidFallback] = &[
    HwidFallback { cur_sku: 125, cur_edition: "EnterpriseS-2021", cur_activation_id: "cce9d2de-98ee-4ce2-8113-222620c64a27", alt_activation_id: "ed655016-a9e8-4434-95d9-4345352c2552", alt_key: "QPM6N-7J2WJ-P88HH-P3YRH-YY74H", alt_edition: "IoTEnterpriseS-2021" },
    HwidFallback { cur_sku: 125, cur_edition: "EnterpriseS-2024", cur_activation_id: "f6e29426-a256-4316-88bf-cc5b0f95ec0c", alt_activation_id: "6c4de1b8-24bb-4c17-9a77-7b939414c298", alt_key: "CGK42-GYN6Y-VD22B-BX98W-J8JXD", alt_edition: "IoTEnterpriseS-2024" },
    HwidFallback { cur_sku: 138, cur_edition: "ProfessionalSingleLanguage", cur_activation_id: "a48938aa-62fa-4966-9d44-9f04da3f72f2", alt_activation_id: "4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c", alt_key: "VK7JG-NPHTM-C97JM-9MPGT-3V66T", alt_edition: "Professional" },
    HwidFallback { cur_sku: 139, cur_edition: "ProfessionalCountrySpecific", cur_activation_id: "f7af7d09-40e4-419c-a49b-eae366689ebd", alt_activation_id: "4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c", alt_key: "VK7JG-NPHTM-C97JM-9MPGT-3V66T", alt_edition: "Professional" },
    HwidFallback { cur_sku: 139, cur_edition: "ProfessionalCountrySpecific-Zn", cur_activation_id: "01eb852c-424d-4060-94b8-c10d799d7364", alt_activation_id: "4de7cb65-cdf1-4de9-8ae8-e3cce27b9f2c", alt_key: "VK7JG-NPHTM-C97JM-9MPGT-3V66T", alt_edition: "Professional" },
];

/// First HWID entry for a SKU id (refine by `version` for multi-branch SKUs).
pub fn entry_for_sku(sku: u32) -> Option<&'static HwidEntry> {
    HWID_KEYS.iter().find(|e| e.sku == sku)
}

/// Fallback mapping for a current SKU + edition, if one exists.
pub fn fallback_for(cur_sku: u32, cur_edition: &str) -> Option<&'static HwidFallback> {
    HWID_FALLBACK
        .iter()
        .find(|f| f.cur_sku == cur_sku && f.cur_edition == cur_edition)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tables_have_the_expected_counts() {
        assert_eq!(HWID_KEYS.len(), 34);
        assert_eq!(HWID_FALLBACK.len(), 5);
    }

    #[test]
    fn every_key_is_a_well_formed_product_key() {
        // 5 groups of 5 uppercase alphanumerics, hyphen-separated.
        let ok = |k: &str| {
            let groups: Vec<&str> = k.split('-').collect();
            groups.len() == 5
                && groups.iter().all(|g| {
                    g.len() == 5 && g.chars().all(|c| c.is_ascii_uppercase() || c.is_ascii_digit())
                })
        };
        assert!(HWID_KEYS.iter().all(|e| ok(e.product_key)), "malformed key");
        assert!(HWID_FALLBACK.iter().all(|f| ok(f.alt_key)), "malformed fallback key");
        // The %f% splitter must be gone.
        assert!(HWID_KEYS.iter().all(|e| !e.product_key.contains('%')));
    }

    #[test]
    fn lookup_resolves_known_editions() {
        assert_eq!(entry_for_sku(48).unwrap().edition_id, "Professional");
        assert_eq!(entry_for_sku(101).unwrap().edition_id, "Core");
        assert!(entry_for_sku(99999).is_none());
    }

    #[test]
    fn fallback_maps_professional_single_language() {
        let f = fallback_for(138, "ProfessionalSingleLanguage").unwrap();
        assert_eq!(f.alt_edition, "Professional");
    }
}
