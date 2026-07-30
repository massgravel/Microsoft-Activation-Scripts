//! `Check_Activation_Status.cmd` — read and present licensing state.
//!
//! This is the most self-contained module: it only *reads* the SPP, so the
//! orchestration and formatting are fully portable and unit-tested here; only
//! the underlying `installed_products` query is Windows-specific.

use crate::error::Result;
use crate::model::Product;
use crate::platform::{LicenseInfo, Spp};

/// A rendered status line for one product.
#[derive(Debug, Clone)]
pub struct StatusLine {
    pub product: Product,
    pub name: String,
    pub partial_key: String,
    pub status: String,
    pub activated: bool,
}

/// Collect status for one product family from the SPP.
pub fn collect(spp: &dyn Spp, product: Product) -> Result<Vec<StatusLine>> {
    let products = spp.installed_products(product)?;
    Ok(products.iter().map(|p| line_for(product, p)).collect())
}

fn line_for(product: Product, info: &LicenseInfo) -> StatusLine {
    StatusLine {
        product,
        name: info.name.clone(),
        partial_key: info
            .partial_product_key
            .clone()
            .unwrap_or_else(|| "—".to_string()),
        status: info.status.label().to_string(),
        activated: info.is_activated(),
    }
}

/// Format collected lines the way MAS prints them: name, partial key, state.
pub fn render(lines: &[StatusLine]) -> String {
    if lines.is_empty() {
        return "No licensed products found.\n".to_string();
    }
    let mut out = String::new();
    for l in lines {
        let mark = if l.activated { "[✓]" } else { "[ ]" };
        out.push_str(&format!(
            "{mark} {} — {}  (…{})\n",
            l.product, l.name, l.partial_key
        ));
        out.push_str(&format!("      {}\n", l.status));
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::LicenseStatus;

    fn info(name: &str, key: Option<&str>, status: LicenseStatus) -> LicenseInfo {
        LicenseInfo {
            activation_id: "id".into(),
            name: name.into(),
            description: "desc".into(),
            partial_product_key: key.map(str::to_string),
            status,
            license_family: None,
            grace_minutes: None,
        }
    }

    #[test]
    fn render_marks_activated_products() {
        let lines = vec![
            line_for(Product::Windows, &info("Pro", Some("ABCDE"), LicenseStatus::Licensed)),
            line_for(Product::Windows, &info("Core", None, LicenseStatus::Notification)),
        ];
        let out = render(&lines);
        assert!(out.contains("[✓] Windows — Pro  (…ABCDE)"));
        assert!(out.contains("[ ] Windows — Core  (…—)"));
        assert!(out.contains("Notification"));
    }

    #[test]
    fn empty_products_render_message() {
        assert_eq!(render(&[]), "No licensed products found.\n");
    }
}
