//! Interactive front-end — the MAS main menu, ported.

pub mod menu;

use crate::activation::{activator_for, status};
use crate::error::Result;
use crate::model::{Method, Product};
use crate::platform::Spp;
use menu::{Menu, MenuItem, Selection};
use std::io::{self, BufRead, Write};

/// One top-level choice from the MAS main menu.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MainChoice {
    Activate(Method, Product),
    CheckStatus,
}

/// Build the main menu, hiding methods that aren't available on this machine
/// (mirrors MAS graying/omitting options by build & product).
pub fn main_menu(windows_build: u32) -> Menu<MainChoice> {
    let mut m = Menu::new("Microsoft Activation Scripts (Rust)");
    let mut key = b'1';
    let mut add = |m: &mut Menu<MainChoice>, label: String, note: &str, choice: MainChoice| {
        m.items.push(
            MenuItem::new(key as char, label, choice).with_note(note.to_string()),
        );
        key += 1;
    };

    for method in Method::ALL {
        for product in [Product::Windows, Product::Office] {
            if method.targets(product) && method.available_for(product, windows_build) {
                add(
                    &mut m,
                    format!("{} — {}", method, product),
                    method.description(),
                    MainChoice::Activate(method, product),
                );
            }
        }
    }
    add(
        &mut m,
        "Check Activation Status".into(),
        "Read-only report for Windows and Office",
        MainChoice::CheckStatus,
    );
    m
}

/// Run one round of the menu against the given I/O and backend. Returns
/// `Ok(true)` to keep looping, `Ok(false)` to exit.
pub fn step<R: BufRead, W: Write>(
    reader: R,
    mut writer: W,
    spp: &dyn Spp,
) -> Result<bool> {
    let build = spp.windows_build().unwrap_or(0);
    let m = main_menu(build);
    write!(writer, "{}", m.render())?;
    let key = menu::prompt_choice(reader, &mut writer, "\nSelect: ")?;

    match m.resolve(key) {
        Selection::Back => Ok(false),
        Selection::Invalid => {
            writeln!(writer, "Invalid choice.")?;
            Ok(true)
        }
        Selection::Value(&choice) => {
            run_choice(&mut writer, spp, choice)?;
            Ok(true)
        }
    }
}

fn run_choice<W: Write>(writer: &mut W, spp: &dyn Spp, choice: MainChoice) -> Result<()> {
    match choice {
        MainChoice::CheckStatus => {
            for product in [Product::Windows, Product::Office] {
                let lines = status::collect(spp, product)?;
                write!(writer, "{}", status::render(&lines))?;
            }
        }
        MainChoice::Activate(method, product) => {
            writeln!(writer, "Running {method} for {product}…")?;
            match activator_for(method).run(spp, product) {
                Ok(()) => writeln!(writer, "  Success.")?,
                Err(e) => writeln!(writer, "  {e}")?,
            }
        }
    }
    Ok(())
}

/// Blocking interactive loop over stdin/stdout.
pub fn run_interactive(spp: &dyn Spp) -> Result<()> {
    let stdin = io::stdin();
    let stdout = io::stdout();
    loop {
        let keep = step(stdin.lock(), stdout.lock(), spp)?;
        if !keep {
            return Ok(());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::platform::stub::StubSpp;

    #[test]
    fn menu_hides_win10_only_methods_on_old_builds() {
        let old = main_menu(7601); // Windows 7
        assert!(!old
            .items
            .iter()
            .any(|it| matches!(it.value, MainChoice::Activate(Method::Hwid, _))));

        let new = main_menu(19045); // Windows 10
        assert!(new
            .items
            .iter()
            .any(|it| matches!(it.value, MainChoice::Activate(Method::Hwid, _))));
    }

    #[test]
    fn check_status_is_always_offered() {
        let m = main_menu(0);
        assert!(m
            .items
            .iter()
            .any(|it| it.value == MainChoice::CheckStatus));
    }

    #[test]
    fn step_check_status_runs_and_loops() {
        // Choose the "Check Activation Status" entry, then verify it loops.
        let m = main_menu(0);
        let status_key = m
            .items
            .iter()
            .find(|it| it.value == MainChoice::CheckStatus)
            .unwrap()
            .key;
        let input = format!("{status_key}\n");
        let mut out = Vec::new();
        let keep = step(input.as_bytes(), &mut out, &StubSpp).unwrap();
        assert!(keep);
        let text = String::from_utf8(out).unwrap();
        assert!(text.contains("No licensed products found."));
    }

    #[test]
    fn step_zero_exits() {
        let mut out = Vec::new();
        let keep = step(&b"0\n"[..], &mut out, &StubSpp).unwrap();
        assert!(!keep);
    }
}
