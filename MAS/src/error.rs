//! Error handling for the MAS port.
//!
//! The original scripts signal failure through `errorlevel` codes, colored
//! console messages and `goto :dk` jumps. Rust replaces all of that with a
//! single typed error carried by [`Result`]. Each variant maps to a concrete
//! failure mode from the batch scripts so callers can react precisely instead
//! of parsing text.
//!
//! Kept dependency-free on purpose (no `thiserror`) so the portable core builds
//! and tests offline on any platform — see the crate root notes.

use std::fmt;

/// Crate-wide result type.
pub type Result<T> = std::result::Result<T, Error>;

/// Everything that can go wrong while inspecting or changing activation state.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// A privileged operation was attempted on a non-Windows build (the stub
    /// Software Protection Platform backend). The real backend requires the
    /// `winapi` feature on a Windows target.
    UnsupportedPlatform {
        /// The operation that is unavailable, e.g. `"install product key"`.
        operation: &'static str,
    },

    /// The process is not running elevated. Almost every MAS operation edits
    /// HKLM / the SPP and needs administrator rights (the scripts self-elevate
    /// via a mshta/PowerShell relaunch).
    NotElevated,

    /// A Windows API / WMI call failed. `hresult` is the raw `HRESULT` where
    /// available (mirrors the scripts surfacing SPP error codes to the user).
    WinApi {
        context: String,
        hresult: Option<i32>,
    },

    /// A required external tool (`ClipUp.exe`, `cscript`, `slmgr.vbs`, `sc.exe`)
    /// could not be found or exited non-zero.
    ExternalTool {
        tool: String,
        detail: String,
    },

    /// The user asked for a product/edition/method that isn't valid for this
    /// machine (e.g. HWID on a build older than Windows 10, or an unknown SKU).
    Unsupported {
        what: String,
    },

    /// The user aborted an interactive prompt.
    Cancelled,

    /// Wrapper for `std::io` failures (file/registry-file/process I/O).
    Io(std::io::Error),
}

impl Error {
    /// Convenience constructor for a WMI/Win32 failure with an `HRESULT`.
    pub fn winapi(context: impl Into<String>, hresult: i32) -> Self {
        Error::WinApi {
            context: context.into(),
            hresult: Some(hresult),
        }
    }

    /// Convenience constructor for a WMI/Win32 failure without a code.
    pub fn winapi_msg(context: impl Into<String>) -> Self {
        Error::WinApi {
            context: context.into(),
            hresult: None,
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::UnsupportedPlatform { operation } => write!(
                f,
                "cannot {operation}: this build has no Windows Software Protection \
                 Platform backend (rebuild on Windows with --features winapi)"
            ),
            Error::NotElevated => {
                write!(f, "administrator privileges are required for this operation")
            }
            Error::WinApi { context, hresult } => match hresult {
                Some(hr) => write!(f, "{context} (HRESULT 0x{:08X})", *hr as u32),
                None => write!(f, "{context}"),
            },
            Error::ExternalTool { tool, detail } => {
                write!(f, "external tool `{tool}` failed: {detail}")
            }
            Error::Unsupported { what } => write!(f, "unsupported: {what}"),
            Error::Cancelled => write!(f, "cancelled by user"),
            Error::Io(e) => write!(f, "I/O error: {e}"),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for Error {
    fn from(e: std::io::Error) -> Self {
        Error::Io(e)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hresult_is_hex_formatted() {
        let e = Error::winapi("activate Windows", 0xC004F074u32 as i32);
        assert_eq!(
            e.to_string(),
            "activate Windows (HRESULT 0xC004F074)"
        );
    }

    #[test]
    fn io_error_chains_source() {
        use std::error::Error as _;
        let e = Error::from(std::io::Error::new(std::io::ErrorKind::NotFound, "x"));
        assert!(e.source().is_some());
    }
}
