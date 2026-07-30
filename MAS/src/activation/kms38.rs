//! KMS38 — extend KMS activation to 2038-01-19.
//!
//! Corrected against the source (`Online_KMS_Activation.cmd`): KMS38 is **not**
//! a ClipUp/gatherosstate flow and not a separate activator. It is a
//! *detect-and-preserve* variant of Online KMS. The 2038 lease is minted by the
//! (emulated) public KMS host replying to a standard KMS-v6 activation for
//! eligible builds; this code only:
//!   1. gates eligibility (build ≥ 14393, not EnterpriseG/GN),
//!   2. installs the GVLK, sets the KMS host, and activates (normal Online KMS),
//!   3. detects the KMS38 lease (`GracePeriodRemaining` > 180 days),
//!   4. pins that product's KMS host to loopback `127.0.0.2:1688` so the renewal
//!      task's public re-activation can't reset it, and skips re-activating.
//!
//! The eligibility/detection/classification logic is pure and tested here; the
//! WMI activation and the loopback registry pin are Windows-only (behind `Spp`).

use crate::activation::Activator;
use crate::data::kms_hosts::{DEFAULT_KMS_HOSTS, DEFAULT_KMS_PORT};
use crate::error::{Error, Result};
use crate::model::{Method, Product};
use crate::platform::Spp;

/// Minimum build for KMS38 (Windows 10 1607 / Server 2016).
pub const MIN_KMS38_BUILD: u32 = 14393;

/// A normal KMS lease is exactly 180 days; anything longer is the 2038 lease.
pub const NORMAL_KMS_GRACE_MINUTES: u32 = 259_200;

/// EnterpriseG / EnterpriseGN activation IDs — explicitly disqualified.
pub const ENTERPRISE_G_IDS: &[&str] = &[
    "e0b2d383-d112-413f-8a80-97f373a5820c", // EnterpriseG  (SKU 171)
    "e38454fb-41a4-4f59-a5dc-25080e354730", // EnterpriseGN (SKU 172)
];

/// The loopback address the KMS38 lock is pinned to.
pub const KMS38_PIN_HOST: &str = "127.0.0.2";

/// Is this product eligible for KMS38?
pub fn is_eligible(build: u32, activation_id: &str) -> bool {
    build >= MIN_KMS38_BUILD
        && !ENTERPRISE_G_IDS
            .iter()
            .any(|id| id.eq_ignore_ascii_case(activation_id))
}

/// Does this grace period indicate an active KMS38 (to-2038) lease?
pub fn is_kms38_lease(build: u32, activation_id: &str, grace_minutes: u32) -> bool {
    is_eligible(build, activation_id) && grace_minutes > NORMAL_KMS_GRACE_MINUTES
}

/// Whole days remaining, rounded up (matches the script's `ceil(gpr/1440)`).
pub fn grace_days(grace_minutes: u32) -> u32 {
    grace_minutes.div_ceil(1440)
}

pub struct Kms38;

impl Activator for Kms38 {
    fn method(&self) -> Method {
        Method::Kms38
    }

    fn run(&self, spp: &dyn Spp, _product: Product) -> Result<()> {
        if !spp.is_elevated() {
            return Err(Error::NotElevated);
        }
        let build = spp.windows_build()?;
        if build < MIN_KMS38_BUILD {
            return Err(Error::Unsupported {
                what: "KMS38 requires Windows 10 1607 / Server 2016 or later".into(),
            });
        }

        // Find an eligible Windows volume product.
        let products = spp.installed_products(Product::Windows)?;
        let target = products
            .into_iter()
            .find(|p| is_eligible(build, &p.activation_id))
            .ok_or_else(|| Error::Unsupported {
                what: "no KMS38-eligible Windows edition installed (EnterpriseG/GN excluded)".into(),
            })?;

        // If already on a 2038 lease, preserve it and stop (skip re-activate).
        if let Some(g) = target.grace_minutes {
            if is_kms38_lease(build, &target.activation_id, g) {
                return spp.pin_kms38(&target.activation_id);
            }
        }

        // Normal Online-KMS activation against the default hosts.
        for host in DEFAULT_KMS_HOSTS {
            if spp.set_kms_host(Product::Windows, host, DEFAULT_KMS_PORT).is_err() {
                continue;
            }
            if spp.activate(Product::Windows, &target.activation_id).is_ok() {
                // Pin the loopback lock so renewal can't reset the lease.
                return spp.pin_kms38(&target.activation_id);
            }
        }
        Err(Error::winapi_msg(
            "KMS38: activation failed against every default host",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn eligibility_gates_build_and_edition() {
        let pro = "00000000-0000-0000-0000-000000000001";
        assert!(is_eligible(14393, pro));
        assert!(!is_eligible(10240, pro)); // too old (1507)
        assert!(!is_eligible(19045, ENTERPRISE_G_IDS[0])); // EnterpriseG excluded
        assert!(!is_eligible(19045, ENTERPRISE_G_IDS[1].to_uppercase().as_str())); // case-insensitive
    }

    #[test]
    fn kms38_lease_detected_only_above_180_days() {
        let pro = "00000000-0000-0000-0000-000000000001";
        assert!(!is_kms38_lease(19045, pro, NORMAL_KMS_GRACE_MINUTES)); // exactly 180d = normal
        assert!(is_kms38_lease(19045, pro, NORMAL_KMS_GRACE_MINUTES + 1));
        // Even a huge grace on EnterpriseG is not a KMS38 lock.
        assert!(!is_kms38_lease(19045, ENTERPRISE_G_IDS[0], 9_000_000));
    }

    #[test]
    fn grace_days_rounds_up() {
        assert_eq!(grace_days(1440), 1);
        assert_eq!(grace_days(1441), 2);
        assert_eq!(grace_days(NORMAL_KMS_GRACE_MINUTES), 180);
    }
}
