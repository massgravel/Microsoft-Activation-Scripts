//! `Separate-Files-Version/Extract_OEM_Folder.cmd` — build a `$OEM$` folder.
//!
//! The original script drops a Windows Setup `$OEM$` payload on the Desktop so
//! that activation runs automatically at OOBE. Windows Setup copies
//! `$OEM$\$$\...` into `%SystemRoot%\...`, so the payload lands at
//! `%SystemRoot%\Setup\Scripts\` and Setup runs `SetupComplete.cmd` there once,
//! post-install. The script's real work is: pick a combination of activators,
//! copy those activator `.cmd` files, and emit the matching `SetupComplete.cmd`.
//!
//! ## Portable (here, unit-tested)
//! * [`BuildOption`] → the exact `SetupComplete.cmd` bytes ([`BuildOption::setup_complete_cmd`]).
//!   This is the heart of the script (the `:*_setup:` blocks + the `:export`
//!   PowerShell that trims them and writes ASCII). Reproduced byte-faithfully,
//!   CRLF line endings, no trailing newline — matching `.Trim()` output.
//! * The `$OEM$\$$\Setup\Scripts` folder layout + which activator files each
//!   option needs ([`BuildOption::activator_files`]).
//!
//! ## Portable but effectful (here, `std::fs`)
//! * [`extract`] creates the folder tree, copies the activator `.cmd` files from
//!   a caller-supplied source dir, and writes `SetupComplete.cmd`. No Windows
//!   API — plain filesystem — so it lives in the module, not behind `Spp`.
//!
//! ## Windows-only (deferred to the caller / a new `Spp` method)
//! * Locating the Desktop (the script reads `User Shell Folders\Desktop` from
//!   the registry, falling back to `%USERPROFILE%\Desktop`). That is a Windows
//!   lookup, so it goes through `Spp::desktop_dir()` (see `new_spp_methods`);
//!   [`extract`] just takes the destination path.
//! * All the batch bootstrapping (self-elevation, x64/ARM relaunch, QuickEdit,
//!   PowerShell/AV sanity checks, update ping) is environment plumbing, not
//!   part of the port.

use crate::error::Result;
use std::io;
use std::path::{Path, PathBuf};

/// Canonical activator script filenames (verbatim from the script's
/// `set <name>=Activators\<name>` table).
const HWID: &str = "HWID_Activation.cmd";
const OHOOK: &str = "Ohook_Activation_AIO.cmd";
const TSFORGE: &str = "TSforge_Activation.cmd";
const KMS: &str = "Online_KMS_Activation.cmd";

/// One `call "%~dp0<file>" <args>` line inside `SetupComplete.cmd`.
type Invocation = (&'static str, &'static str);

/// The seven menu combinations the script offers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuildOption {
    /// `[1]` HWID — Windows.
    Hwid,
    /// `[2]` Ohook — Office.
    Ohook,
    /// `[3]` TSforge — Windows / ESU / Office.
    TSforge,
    /// `[4]` Online KMS — Windows / Office.
    Kms,
    /// `[5]` HWID + Ohook.
    HwidOhook,
    /// `[6]` HWID + Ohook + TSforge (ESU).
    HwidOhookTsforge,
    /// `[7]` TSforge (Windows / ESU) + Ohook.
    TsforgeOhook,
}

impl BuildOption {
    /// All options in menu order.
    pub const ALL: [BuildOption; 7] = [
        BuildOption::Hwid,
        BuildOption::Ohook,
        BuildOption::TSforge,
        BuildOption::Kms,
        BuildOption::HwidOhook,
        BuildOption::HwidOhookTsforge,
        BuildOption::TsforgeOhook,
    ];

    /// The menu hotkey digit (`choice /C:1234567R0`).
    pub fn key(self) -> char {
        match self {
            BuildOption::Hwid => '1',
            BuildOption::Ohook => '2',
            BuildOption::TSforge => '3',
            BuildOption::Kms => '4',
            BuildOption::HwidOhook => '5',
            BuildOption::HwidOhookTsforge => '6',
            BuildOption::TsforgeOhook => '7',
        }
    }

    /// The success label (verbatim from the script's `set oem=` lines).
    pub fn label(self) -> &'static str {
        match self {
            BuildOption::Hwid => "HWID",
            BuildOption::Ohook => "Ohook",
            BuildOption::TSforge => "TSforge",
            BuildOption::Kms => "Online KMS",
            BuildOption::HwidOhook => "HWID [Windows] + Ohook [Office]",
            BuildOption::HwidOhookTsforge => "HWID [Windows] + Ohook [Office] + TSforge [ESU]",
            BuildOption::TsforgeOhook => "TSforge [Windows / ESU] + Ohook [Office]",
        }
    }

    /// The activator calls this option runs, in order, as `(filename, args)`.
    /// Args are transcribed verbatim from the `:*_setup:` blocks.
    pub fn invocations(self) -> Vec<Invocation> {
        match self {
            BuildOption::Hwid => vec![(HWID, "/HWID")],
            BuildOption::Ohook => vec![(OHOOK, "/Ohook")],
            BuildOption::TSforge => vec![(TSFORGE, "/Z-WindowsESUOffice")],
            BuildOption::Kms => vec![(KMS, "/K-WindowsOffice")],
            BuildOption::HwidOhook => vec![(HWID, "/HWID"), (OHOOK, "/Ohook")],
            BuildOption::HwidOhookTsforge => {
                vec![(HWID, "/HWID"), (OHOOK, "/Ohook"), (TSFORGE, "/Z-ESU")]
            }
            BuildOption::TsforgeOhook => {
                vec![(TSFORGE, "/Z-Windows /Z-ESU"), (OHOOK, "/Ohook")]
            }
        }
    }

    /// Distinct activator files this option copies into the payload, in order.
    pub fn activator_files(self) -> Vec<&'static str> {
        let mut files = Vec::new();
        for (f, _) in self.invocations() {
            if !files.contains(&f) {
                files.push(f);
            }
        }
        files
    }

    /// The exact `SetupComplete.cmd` text for this option.
    ///
    /// Byte-faithful to what the batch `:export` step produced: the trimmed
    /// `:*_setup:` block, CRLF line endings, ASCII, no trailing newline. A
    /// single activator is called directly; multiple are each wrapped in a
    /// `setlocal`/`endlocal` pair (as the combo blocks do).
    pub fn setup_complete_cmd(self) -> String {
        let mut lines: Vec<String> = vec![
            "@echo off".into(),
            String::new(),
            "fltmc >nul || exit /b".into(),
            String::new(),
        ];

        let invs = self.invocations();
        if invs.len() == 1 {
            let (file, args) = invs[0];
            lines.push(format!("call \"%~dp0{file}\" {args}"));
        } else {
            for (i, (file, args)) in invs.iter().enumerate() {
                if i > 0 {
                    lines.push(String::new());
                }
                lines.push("setlocal".into());
                lines.push(format!("call \"%~dp0{file}\" {args}"));
                lines.push("endlocal".into());
            }
        }

        lines.push(String::new());
        lines.push("cd \\".into());
        lines.push(
            r#"(goto) 2>nul & (if "%~dp0"=="%SystemRoot%\Setup\Scripts\" rd /s /q "%~dp0")"#.into(),
        );

        lines.join("\r\n")
    }
}

/// Relative payload path under the Desktop: `$OEM$\$$\Setup\Scripts`.
///
/// Windows Setup strips the leading `$OEM$\$$` and copies the rest into
/// `%SystemRoot%`, so this ends up at `%SystemRoot%\Setup\Scripts`.
fn scripts_dir(dest_root: &Path) -> PathBuf {
    dest_root
        .join("$OEM$")
        .join("$$")
        .join("Setup")
        .join("Scripts")
}

/// Build the `$OEM$` payload under `dest_root` (the Desktop).
///
/// `activators_src` is the folder holding the activator `.cmd` files (the
/// script's `Activators\` directory). Returns the created `...\Setup\Scripts`
/// directory. Mirrors the script's guards: refuse if `$OEM$` already exists,
/// and fail up front if a required activator file is missing.
pub fn extract(option: BuildOption, activators_src: &Path, dest_root: &Path) -> Result<PathBuf> {
    let oem_root = dest_root.join("$OEM$");
    if oem_root.exists() {
        return Err(io::Error::new(
            io::ErrorKind::AlreadyExists,
            "The $OEM$ folder already exists on your Desktop.",
        )
        .into());
    }

    // Faithful to the `_nofile` pre-check: verify sources before touching disk.
    let files = option.activator_files();
    for f in &files {
        let src = activators_src.join(f);
        if !src.is_file() {
            return Err(io::Error::new(
                io::ErrorKind::NotFound,
                format!("Missing file in the 'Activators' folder: {f}"),
            )
            .into());
        }
    }

    let scripts = scripts_dir(dest_root);
    std::fs::create_dir_all(&scripts).map_err(crate::error::Error::Io)?;

    for f in &files {
        std::fs::copy(activators_src.join(f), scripts.join(f)).map_err(crate::error::Error::Io)?;
    }

    // ASCII by construction, so bytes == what `WriteAllText(..., ASCII)` emitted.
    std::fs::write(scripts.join("SetupComplete.cmd"), option.setup_complete_cmd())
        .map_err(crate::error::Error::Io)?;

    Ok(scripts)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hwid_setup_is_byte_faithful() {
        let expected = "@echo off\r\n\
            \r\n\
            fltmc >nul || exit /b\r\n\
            \r\n\
            call \"%~dp0HWID_Activation.cmd\" /HWID\r\n\
            \r\n\
            cd \\\r\n\
            (goto) 2>nul & (if \"%~dp0\"==\"%SystemRoot%\\Setup\\Scripts\\\" rd /s /q \"%~dp0\")";
        assert_eq!(BuildOption::Hwid.setup_complete_cmd(), expected);
    }

    #[test]
    fn tsforge_office_uses_combined_switch() {
        // Single-activator TSforge takes the combined /Z-WindowsESUOffice switch.
        assert!(BuildOption::TSforge
            .setup_complete_cmd()
            .contains("call \"%~dp0TSforge_Activation.cmd\" /Z-WindowsESUOffice"));
    }

    #[test]
    fn tsforge_ohook_setup_is_byte_faithful() {
        let expected = "@echo off\r\n\
            \r\n\
            fltmc >nul || exit /b\r\n\
            \r\n\
            setlocal\r\n\
            call \"%~dp0TSforge_Activation.cmd\" /Z-Windows /Z-ESU\r\n\
            endlocal\r\n\
            \r\n\
            setlocal\r\n\
            call \"%~dp0Ohook_Activation_AIO.cmd\" /Ohook\r\n\
            endlocal\r\n\
            \r\n\
            cd \\\r\n\
            (goto) 2>nul & (if \"%~dp0\"==\"%SystemRoot%\\Setup\\Scripts\\\" rd /s /q \"%~dp0\")";
        assert_eq!(BuildOption::TsforgeOhook.setup_complete_cmd(), expected);
    }

    #[test]
    fn combo_blocks_wrap_every_call_in_setlocal() {
        let text = BuildOption::HwidOhookTsforge.setup_complete_cmd();
        assert_eq!(text.matches("setlocal\r\n").count(), 3);
        assert_eq!(text.matches("\r\nendlocal").count(), 3);
        assert!(text.contains("call \"%~dp0TSforge_Activation.cmd\" /Z-ESU"));
    }

    #[test]
    fn output_is_ascii_and_has_no_trailing_newline() {
        for opt in BuildOption::ALL {
            let s = opt.setup_complete_cmd();
            assert!(s.is_ascii(), "{:?} not ASCII", opt);
            assert!(!s.ends_with('\n'), "{:?} has trailing newline", opt);
            assert!(s.starts_with("@echo off\r\n"));
        }
    }

    #[test]
    fn activator_files_are_distinct_and_ordered() {
        assert_eq!(BuildOption::Hwid.activator_files(), vec![HWID]);
        assert_eq!(
            BuildOption::HwidOhookTsforge.activator_files(),
            vec![HWID, OHOOK, TSFORGE]
        );
        // Ohook appears once even though it is only ever one call.
        assert_eq!(BuildOption::TsforgeOhook.activator_files(), vec![TSFORGE, OHOOK]);
    }

    #[test]
    fn keys_match_the_menu() {
        let keys: Vec<char> = BuildOption::ALL.iter().map(|o| o.key()).collect();
        assert_eq!(keys, vec!['1', '2', '3', '4', '5', '6', '7']);
    }
}
