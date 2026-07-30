//! Default KMS host list for Online KMS activation.
//!
//! MAS lets the user pass a server with `/K-Server-…` and otherwise tries a
//! built-in list of public KMS hosts. Ported as data, not code.
//!
//! ponytail: static default list — the real tuning knob. Public KMS hosts come
//! and go; expose `--kms-host` (already supported via [`OnlineKms`]) and refresh
//! this list rather than hard-depending on any single entry.

/// Standard KMS port.
pub const DEFAULT_KMS_PORT: u16 = 1688;

/// Fallback public KMS hosts, tried in order until one activates.
pub const DEFAULT_KMS_HOSTS: &[&str] = &[
    "kms.digiboy.ir",
    "kms8.msguides.com",
    "kms.03k.org",
    "kms.chinancce.com",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn host_list_is_nonempty_and_port_is_standard() {
        assert!(!DEFAULT_KMS_HOSTS.is_empty());
        assert_eq!(DEFAULT_KMS_PORT, 1688);
    }
}
