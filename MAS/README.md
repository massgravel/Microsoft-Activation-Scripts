# MAS — Rust port

An idiomatic Rust re-architecture of the Microsoft Activation Scripts. This is a
**re-architecture, not a line-by-line transliteration**: the batch scripts'
`goto`/`errorlevel`/global-`set` style is replaced with typed enums, `Result`
error flow, and a single trait boundary between portable logic and Windows
internals.

## Architecture

```
src/
  model.rs              Product / Method / LicenseStatus enums + the real SPP GUIDs
  error.rs              typed Error/Result (replaces errorlevel + colored echoes)
  data/kms_hosts.rs     ported data tables
  platform/
    mod.rs              `Spp` trait — the ONE seam to Windows; safe API over FFI
    stub.rs             non-Windows backend (read-only no-ops; privileged ops error)
    windows_backend.rs  real WMI/COM backend (cfg(windows) + feature "winapi")
  activation/
    status.rs            Check_Activation_Status.cmd  [ported: read path]
    online_kms.rs        Online_KMS_Activation.cmd    [ported: orchestration]
    hwid.rs / kms38.rs   HWID_Activation.cmd          [scaffold + spec]
    ohook.rs             Ohook_Activation_AIO.cmd     [scaffold + spec]
    tsforge.rs           TSforge_Activation.cmd       [scaffold + spec]
  cli/                  the interactive menu (MAS_AIO front-end)
```

Every Windows effect — WMI calls, ClipUp, registry — lives behind
[`platform::Spp`]. The rest of the crate is OS-independent and unit-tested.

## Build & test

```sh
# Portable core — builds and tests on any OS, zero dependencies:
cargo test

# Real Windows backend (WMI/COM via the `windows` crate):
cargo build --release --features winapi --target x86_64-pc-windows-msvc
```

The Windows backend is written against `windows` 0.58 from the documented SPP
WMI contract. It has **not** been compiled on the porting host (Linux has no
Windows target) — expect to shake out minor `windows`-crate signature drift on a
real Windows build. The portable core (menus, model, data, orchestration) is
compiled and tested (`20 passing`).

## Port status & roadmap

The three deepest activators are **not fakeable**; each carries its analyzed spec
in the module docs as the roadmap:

| Module   | Why it's hard (from analysis)                                                        |
|----------|-------------------------------------------------------------------------------------|
| TSforge  | Byte-exact reverse-engineered SPP trusted-store (`data.dat`: RSA/AES/HMAC, per-block CRC32/SHA-256, per-OS alignment). A faithful port ≈ porting **LibTSforge** to a pure `libtsforge` crate + thin FFI to swap `data.dat`. |
| Ohook    | Ships two reverse-engineered SPP-client DLLs (`sppc32/64.dll`, pinned SHA-256) that forward to the renamed genuine DLL. The blobs are **assets**, not code; the portable planner + license install is the port. |
| status   | The rich detail uses the **undocumented SLC private ABI** (`SLGet*Information`, 40-byte struct stride, `KUSER_SHARED_DATA @0x7FFE02C8`). The WMI backend here covers the common case; the SLC path is a future direct-FFI slice. |

`HWID`/`KMS38` are a shorter hop: generate `GenuineTicket.xml` (portable) and
delegate to `ClipUp.exe` + ClipSVC (the script does the same — the ticket
consumption is closed Windows behavior).

## What is deliberately faithful

- SPP `ApplicationID` GUIDs, WMI class names, method + parameter names, and
  `LicenseStatus` codes are the real Microsoft values (see `model.rs` tests).
- Online KMS falls through a host list exactly like the script; `--kms-host`
  overrides it.
- Nothing reports "activated" that didn't actually activate — unported paths
  return a typed `Unsupported` error instead of lying.
