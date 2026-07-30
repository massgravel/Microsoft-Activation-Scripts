//! Terminal menu rendering and selection.
//!
//! The scripts build their colored menus with hundreds of `echo` lines and
//! parse the choice with `choice`/`set /p` plus a wall of `if "%o%"=="1"`.
//! This is the portable, testable replacement: a menu is data, rendering is a
//! pure function, and selection maps a keypress back to a typed value.

use std::fmt;
use std::io::{self, Write};

/// A selectable menu entry associating a hotkey with a payload `T`.
pub struct MenuItem<T> {
    pub key: char,
    pub label: String,
    /// Optional dim secondary line (method description, availability note).
    pub note: Option<String>,
    pub value: T,
}

impl<T> MenuItem<T> {
    pub fn new(key: char, label: impl Into<String>, value: T) -> Self {
        MenuItem {
            key,
            label: label.into(),
            note: None,
            value,
        }
    }

    pub fn with_note(mut self, note: impl Into<String>) -> Self {
        self.note = Some(note.into());
        self
    }
}

/// A titled list of items plus an implicit `0` = back/exit.
pub struct Menu<T> {
    pub title: String,
    pub items: Vec<MenuItem<T>>,
}

impl<T> Menu<T> {
    pub fn new(title: impl Into<String>) -> Self {
        Menu {
            title: title.into(),
            items: Vec::new(),
        }
    }

    pub fn item(mut self, item: MenuItem<T>) -> Self {
        self.items.push(item);
        self
    }

    /// Render the menu to a string. Pure — no I/O — so it can be asserted in
    /// tests. `0) Exit` is always appended, matching MAS's convention.
    pub fn render(&self) -> String {
        let mut out = String::new();
        out.push_str(&self.title);
        out.push('\n');
        out.push_str(&"=".repeat(self.title.len()));
        out.push('\n');
        for it in &self.items {
            out.push_str(&format!("  {}) {}\n", it.key, it.label));
            if let Some(note) = &it.note {
                out.push_str(&format!("       {note}\n"));
            }
        }
        out.push_str("  0) Back / Exit\n");
        out
    }

    /// Resolve a keypress to the item's value. `'0'` yields `None` (back/exit).
    pub fn resolve(&self, key: char) -> Selection<&T> {
        if key == '0' {
            return Selection::Back;
        }
        match self
            .items
            .iter()
            .find(|it| it.key.eq_ignore_ascii_case(&key))
        {
            Some(it) => Selection::Value(&it.value),
            None => Selection::Invalid,
        }
    }
}

/// Result of matching a keypress against a [`Menu`].
#[derive(Debug, PartialEq, Eq)]
pub enum Selection<T> {
    Value(T),
    Back,
    Invalid,
}

/// Prompt on the given writer, read one line from the reader, and return the
/// trimmed first character. Split from [`Menu`] so tests drive it with cursors.
pub fn prompt_choice<R: io::BufRead, W: Write>(
    mut reader: R,
    mut writer: W,
    prompt: &str,
) -> io::Result<char> {
    write!(writer, "{prompt}")?;
    writer.flush()?;
    let mut line = String::new();
    reader.read_line(&mut line)?;
    Ok(line.trim().chars().next().unwrap_or('\0'))
}

impl<T> fmt::Display for Menu<T> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.render())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> Menu<u8> {
        Menu::new("Activate")
            .item(MenuItem::new('1', "HWID", 1u8).with_note("permanent"))
            .item(MenuItem::new('2', "Online KMS", 2u8))
    }

    #[test]
    fn render_lists_items_and_exit() {
        let r = sample().render();
        assert!(r.contains("1) HWID"));
        assert!(r.contains("permanent"));
        assert!(r.contains("0) Back / Exit"));
    }

    #[test]
    fn resolve_maps_keys_and_back() {
        let m = sample();
        assert_eq!(m.resolve('1'), Selection::Value(&1u8));
        assert_eq!(m.resolve('0'), Selection::Back);
        assert_eq!(m.resolve('9'), Selection::Invalid);
    }

    #[test]
    fn prompt_reads_first_char() {
        let input = b"2\n";
        let mut out = Vec::new();
        let c = prompt_choice(&input[..], &mut out, "> ").unwrap();
        assert_eq!(c, '2');
        assert_eq!(out, b"> ");
    }
}
