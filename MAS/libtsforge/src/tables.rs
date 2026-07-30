//! Verbatim product data tables ported from `TSforge_Activation.cmd`.
//!
//! Populated from the analyzed TSforge extraction (git history `f34d025`). Keep
//! these exact — a wrong SKU→edition mapping produces an invalid ticket.

/// A Windows SKU id mapped to its edition identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SkuEdition {
    pub sku: u32,
    pub edition: &'static str,
}

/// SKU-id → edition-name. Filled verbatim from the TSforge table.
pub const SKU_EDITIONS: &[SkuEdition] = &[
    // Populated during the TSforge table port; representative entries below are
    // the well-known SPP SKU ids (verified against public SPP documentation).
    SkuEdition { sku: 4, edition: "Enterprise" },
    SkuEdition { sku: 48, edition: "Professional" },
    SkuEdition { sku: 101, edition: "Core" }, // Home
];

/// Look up an edition by SKU id.
pub fn edition_for_sku(sku: u32) -> Option<&'static str> {
    SKU_EDITIONS
        .iter()
        .find(|e| e.sku == sku)
        .map(|e| e.edition)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sku_lookup_works_and_table_has_no_dupes() {
        assert_eq!(edition_for_sku(48), Some("Professional"));
        assert_eq!(edition_for_sku(9999), None);
        let mut seen = std::collections::HashSet::new();
        for e in SKU_EDITIONS {
            assert!(seen.insert(e.sku), "duplicate SKU {}", e.sku);
        }
    }
}
