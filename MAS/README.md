# MAS — Rust port

An idiomatic Rust re-architecture of the Microsoft Activation Scripts. This is a
**re-architecture, not a line-by-line transliteration**: the batch scripts'
`goto`/`errorlevel`/global-`set` style is replaced with typed enums, `Result`
error flow, and a single trait boundary between portable logic and Windows
internals.

## Architecture

```
MAS/
  src/
    model.rs              Product / Method / LicenseStatus enums + the real SPP GUIDs
    error.rs              typed Error/Result (replaces errorlevel + colored echoes)
    data/kms_hosts.rs     ported data tables
    platform/
      mod.rs              `Spp` trait — the ONE seam to Windows; safe API over FFI
      stub.rs             non-Windows backend (read-only no-ops; privileged ops error)
      windows_backend.rs  real WMI/COM + ClipSVC backend (cfg(windows) + "winapi")
      test_util.rs        in-memory Spp fake for unit tests
    activation/
      status.rs            Check_Activation_Status.cmd  [ported: read path]
      online_kms.rs        Online_KMS_Activation.cmd    [ported: orchestration]
      hwid.rs              HWID_Activation.cmd          [ported: region + apply]
      kms38.rs             (KMS38 sub-flow)             [ported: detect/preserve]
      ohook.rs             Ohook_Activation_AIO.cmd     [scaffold + spec]
      tsforge.rs           TSforge_Activation.cmd       [wired to libtsforge]
    cli/                  the interactive menu (MAS_AIO front-end)
  libtsforge/            SPP trusted-store codec (own crate, dependency-free)
    src/crc32.rs           CRC-32/BZIP2  (test-vector verified)
    src/sha256.rs          SHA-256       (test-vector verified)
    src/common.rs          PsVersion detection, Align, UTF-16, block/AES constants
    src/physical_store.rs  Vista / Win7 block dialects (round-trip tested)
    src/variable_bag.rs    the two CRC-block dialects (round-trip + CRC tested)
    src/{store,tables,crypto}.rs   error type, product tables, crypto trait seam
```

Every Windows effect — WMI, ClipUp/ClipSVC, registry — lives behind
[`platform::Spp`]. The rest of the crate is OS-independent and unit-tested
(**48 tests**: 31 in `mas`, 17 in `libtsforge`), zero dependencies, offline.

## Build & test

```sh
cd MAS
cargo test                    # 48 passing — portable core, any OS, no Windows required
cargo clippy --workspace       # clean

# Real Windows backend (WMI/COM via the `windows` crate):
cargo build --release --features winapi --target x86_64-pc-windows-msvc
```

The Windows backend is written against `windows` 0.58 from the documented SPP
WMI contract plus `std::process`/`std::fs` for ClipUp/ClipSVC/`reg`. It has
**not** been compiled on the porting host (Linux has no Windows target) — expect
to shake out minor signature drift on a real Windows build.

## Port status & roadmap

Ported and tested (pure logic — the Windows effects sit behind `Spp`):

| Module     | What's ported                                                                        |
|------------|--------------------------------------------------------------------------------------|
| status     | Read path: SPP query → typed `LicenseStatus` → formatted report.                     |
| online_kms | GVLK install → KMS host fallback loop → activate.                                    |
| hwid       | Region decision (30-country skip → GeoId 244) + two-method apply (ClipSVC restart → `clipup -v -o`, success = `tokens.dat`). |
| kms38      | Eligibility gate (build ≥ 14393, EnterpriseG/GN excluded), KMS38-lease detection (>180 days), loopback `127.0.0.2` pin, skip-reactivate. |
| libtsforge | CRC-32/BZIP2, SHA-256, `PsVersion` detection, alignment, UTF-16, the Vista/Win7 physical-store dialects and both VariableBag CRC dialects — all round-trip/vector tested. |

Remaining depth (each carries its analyzed spec in the module docs):

| Module   | Why it's hard (from analysis)                                                        |
|----------|-------------------------------------------------------------------------------------|
| TSforge  | Byte-exact SPP trusted store. `libtsforge` ports the readable LibTSforge C# source; still needs `TokenStoreModern`, the Modern physical store, the RSA CryptoAPI-blob crypto (behind a `crypto` feature), and the verbatim KMS/HWID response blobs. |
| Ohook    | Ships two reverse-engineered SPP-client DLLs (`sppc32/64.dll`, pinned SHA-256) that forward to the renamed genuine DLL. The blobs are **assets**, not code; the portable planner + license install is the port. |
| status   | The rich detail uses the **undocumented SLC private ABI** (`SLGet*Information`, struct stride, `KUSER_SHARED_DATA @0x7FFE02C8`). The WMI backend covers the common case. |
| HWID gen | `GenuineTicket.xml` is a pure RSA-signed XML build (embedded 1024-bit `clientLockboxKey`, same CryptoAPI blob format as TSforge). Portable once the shared RSA-blob layer lands; today generation sits behind `Spp`. |

## What is deliberately faithful

- SPP `ApplicationID` GUIDs, WMI class names, method + parameter names, and
  `LicenseStatus` codes are the real Microsoft values (see `model.rs` tests).
- CRC-32 is the **non-reflected BZIP2** variant the store actually uses
  (verified: `"123456789"` → `0xFC891918`), not the reflected IEEE CRC.
- KMS38 is a detect-and-preserve variant of Online KMS (not a ClipUp flow); the
  `127.0.0.2` pin is the on-disk signature of the lock, exactly as the script.
- Nothing reports "activated" that didn't actually activate — unported paths
  return a typed `Unsupported` error instead of lying.
