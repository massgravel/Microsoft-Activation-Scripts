//! Core domain model.
//!
//! The batch scripts encode everything as bare strings and magic GUIDs sprayed
//! across thousands of lines. Here that becomes a small set of enums with the
//! constants attached to them, so the compiler enforces which activation method
//! is valid for which product and there is exactly one source of truth for the
//! Software Protection Platform application IDs.

use std::fmt;

/// A licensable Microsoft product family.
///
/// Windows and Office are handled by two different Software Protection Platform
/// providers, each identified by its own `ApplicationID` GUID and exposed
/// through different WMI classes. Encoding that here removes the `if office`
/// branching that is duplicated all over the scripts.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Product {
    Windows,
    Office,
}

impl Product {
    /// The SPP `ApplicationID` GUID used in every WMI query
    /// (`WHERE ApplicationID='…'`).
    pub const fn application_id(self) -> &'static str {
        match self {
            Product::Windows => "55c92734-d682-4d71-983e-d6ec3f16059f",
            Product::Office => "0ff1ce15-a989-479d-af46-f275c6370663",
        }
    }

    /// WMI class exposing the per-product licensing state
    /// (`Get PartialProductKey`, `LicenseStatus`, …).
    pub const fn product_class(self) -> &'static str {
        match self {
            Product::Windows => "SoftwareLicensingProduct",
            // Office ≤2016 with the C2R/MSI SPP uses the OSPP class; MAS falls
            // back to SoftwareLicensingProduct on newer Office builds.
            Product::Office => "OfficeSoftwareProtectionProduct",
        }
    }

    /// WMI class exposing the service-wide operations
    /// (`InstallProductKey`, `SetKeyManagementServiceMachine`, …).
    pub const fn service_class(self) -> &'static str {
        match self {
            Product::Windows => "SoftwareLicensingService",
            Product::Office => "OfficeSoftwareProtectionService",
        }
    }

    pub const fn label(self) -> &'static str {
        match self {
            Product::Windows => "Windows",
            Product::Office => "Office",
        }
    }
}

impl fmt::Display for Product {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.label())
    }
}

/// An activation method offered by MAS.
///
/// Each variant carries the rules that the scripts check inline before letting
/// you pick it: which product it targets and the minimum Windows build it needs.
/// [`Method::available_for`] centralises those guards.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Method {
    /// Hardware-ID / digital-license activation (permanent, tied to the device).
    Hwid,
    /// KMS38 — extends the KMS activation expiry to the year 2038 via ClipUp.
    Kms38,
    /// Classic KMS client: install GVLK, point at a KMS host, activate.
    OnlineKms,
    /// Office activation by hooking the SPP (`sppc`/`osppc`) — no key install.
    Ohook,
    /// TSforge — writes activation tickets straight into the SPP token store.
    TSforge,
}

impl Method {
    pub const ALL: [Method; 5] = [
        Method::Hwid,
        Method::Kms38,
        Method::OnlineKms,
        Method::Ohook,
        Method::TSforge,
    ];

    pub const fn label(self) -> &'static str {
        match self {
            Method::Hwid => "HWID",
            Method::Kms38 => "KMS38",
            Method::OnlineKms => "Online KMS",
            Method::Ohook => "Ohook",
            Method::TSforge => "TSforge",
        }
    }

    /// One-line description shown in the menu.
    pub const fn description(self) -> &'static str {
        match self {
            Method::Hwid => "Permanent activation for Windows 10/11 (digital license)",
            Method::Kms38 => "Activate Windows until 2038-01-19 via KMS",
            Method::OnlineKms => "180-day renewable KMS activation for Windows and Office",
            Method::Ohook => "Permanent Office activation (no key, hooks the SPP)",
            Method::TSforge => "Ticket-based activation for Windows, Office and ESU",
        }
    }

    /// Whether this method can act on the given product at all.
    pub const fn targets(self, product: Product) -> bool {
        match self {
            // Windows-only digital-license / KMS-expiry tricks.
            Method::Hwid | Method::Kms38 => matches!(product, Product::Windows),
            // Office-only SPP hook.
            Method::Ohook => matches!(product, Product::Office),
            // These handle both families.
            Method::OnlineKms | Method::TSforge => true,
        }
    }

    /// Minimum Windows build required, if any. HWID and KMS38 rely on
    /// `ClipUp.exe`/digital-license infrastructure introduced in Windows 10
    /// (build 10240).
    pub const fn min_windows_build(self) -> Option<u32> {
        match self {
            Method::Hwid | Method::Kms38 => Some(10240),
            _ => None,
        }
    }

    /// Full guard combining product targeting and the OS build check — the
    /// scripts scatter this logic; here it is one call.
    pub fn available_for(self, product: Product, windows_build: u32) -> bool {
        if !self.targets(product) {
            return false;
        }
        match self.min_windows_build() {
            Some(min) => windows_build >= min,
            None => true,
        }
    }
}

impl fmt::Display for Method {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.label())
    }
}

/// Licensing state of a single product, as read from the SPP.
///
/// Mirrors the WMI `LicenseStatus` enumeration used throughout
/// `Check_Activation_Status.cmd`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LicenseStatus {
    Unlicensed,
    Licensed,
    OobGrace,
    OotGrace,
    NonGenuineGrace,
    Notification,
    ExtendedGrace,
}

impl LicenseStatus {
    /// Map the raw WMI integer to the enum. Values are Microsoft's, quoted
    /// verbatim in the scripts.
    pub const fn from_wmi(value: u32) -> Option<Self> {
        Some(match value {
            0 => LicenseStatus::Unlicensed,
            1 => LicenseStatus::Licensed,
            2 => LicenseStatus::OobGrace,
            3 => LicenseStatus::OotGrace,
            4 => LicenseStatus::NonGenuineGrace,
            5 => LicenseStatus::Notification,
            6 => LicenseStatus::ExtendedGrace,
            _ => return None,
        })
    }

    pub const fn label(self) -> &'static str {
        match self {
            LicenseStatus::Unlicensed => "Unlicensed",
            LicenseStatus::Licensed => "Licensed",
            LicenseStatus::OobGrace => "Initial grace period",
            LicenseStatus::OotGrace => "Additional grace period (KMS expired)",
            LicenseStatus::NonGenuineGrace => "Non-genuine grace period",
            LicenseStatus::Notification => "Notification",
            LicenseStatus::ExtendedGrace => "Extended grace period",
        }
    }

    /// Is the product currently activated?
    pub const fn is_activated(self) -> bool {
        matches!(self, LicenseStatus::Licensed)
    }
}

impl fmt::Display for LicenseStatus {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.label())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hwid_is_windows_only_and_win10_plus() {
        assert!(Method::Hwid.available_for(Product::Windows, 19045));
        assert!(!Method::Hwid.available_for(Product::Windows, 9600)); // Win 8.1
        assert!(!Method::Hwid.available_for(Product::Office, 22631)); // wrong product
    }

    #[test]
    fn ohook_is_office_only() {
        assert!(Method::Ohook.targets(Product::Office));
        assert!(!Method::Ohook.targets(Product::Windows));
    }

    #[test]
    fn online_kms_targets_both() {
        assert!(Method::OnlineKms.available_for(Product::Windows, 7601));
        assert!(Method::OnlineKms.available_for(Product::Office, 7601));
    }

    #[test]
    fn application_ids_are_the_real_spp_guids() {
        assert_eq!(
            Product::Windows.application_id(),
            "55c92734-d682-4d71-983e-d6ec3f16059f"
        );
        assert_eq!(
            Product::Office.application_id(),
            "0ff1ce15-a989-479d-af46-f275c6370663"
        );
    }

    #[test]
    fn license_status_round_trips_known_values() {
        assert_eq!(LicenseStatus::from_wmi(1), Some(LicenseStatus::Licensed));
        assert!(LicenseStatus::from_wmi(1).unwrap().is_activated());
        assert_eq!(LicenseStatus::from_wmi(99), None);
    }
}
