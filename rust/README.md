# xq-rs — Rust reimplementation of xq

A small CLI that reads `.xlsx` / `.xlsm` / `.xlsb` / `.xls` and emits JSON/NDJSON
for `jq`. It uses [`calamine`](https://crates.io/crates/calamine) for parsing and
`serde_json` for output.

## Build

Native (Windows / Linux / macOS):

```bash
cargo build --release
```

Windows cross-build from Alpine (what was used for the checked-in `xq-rs.exe`):

```bash
apk add mingw-w64-gcc
rustup target add x86_64-pc-windows-gnu
cargo build --release --target x86_64-pc-windows-gnu
```

The resulting binary is
`target/x86_64-pc-windows-gnu/release/xq-rs.exe` (or `target/release/xq-rs`).

## Usage

```text
xq-rs sheets <file>
xq-rs cells  <file> [--sheet NAME] [--range A1:C10|NAME] [--json] [--formulas] [--all] [--text] [--limit N]
xq-rs names  <file> [--json]
xq-rs vba    <file> [--json]
xq-rs info   <file>
```

Examples:

```bash
xq-rs cells book.xlsx --formulas --json
xq-rs cells book.xlsb --range A1:C10 --json
xq-rs cells book.xls  --sheet Sheet1 --text
xq-rs names book.xlsx --json
xq-rs vba   book.xlsm --json
xq-rs info  book.xls
```

## Differences from the C# version

- `--text` only emits the raw value as text; Excel `numFmt` formatting is not
  implemented in this version.
- `names` always reports `scope:"workbook"` because `calamine` does not expose
  local-sheet scope / hidden / comment metadata.
- `vba` modules contain `name` and `code` only (no `kind`, `stream`, `offset`,
  `readonly`, `private`).
- `calamine`'s `.xls` formula decompiler can emit `$USM$` for some 3D references;
  `.xlsb` external-workbook references may appear as `EXTERNAL_WB_NAME`.
- Values are emitted as JSON numbers (e.g. `1.0`) rather than the C# version's
  compact integers (`1`); both are valid JSON.
- `--limit N` stops after N emitted cells. The C# version stops reading early;
  this Rust version still parses the whole sheet via `calamine` and only limits
  the output.

## Benchmark

`../bench.ps1` compares this binary with the C# `xq.exe` and an Excel COM reader.
See the top-level `README.md` for results.
