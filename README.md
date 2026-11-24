# xq — xlsx / xlsm / xlsb / xls を jq に流すための C# / .NET 10 CLI

`xq` は **.xlsx / .xlsm / .xlsb / .xls のセル・数式・名前定義・VBA を JSON で出力する単体実行ファイル**です。
Excel も COM も使わず、ZIP + XML + BIFF + OLE を直接読むので高速です。

構成:

```
xq/
├── src/          C# ソース
├── xq.csproj     .NET 10 プロジェクト
├── tests/        xUnit テスト + CLI スモークテスト
├── build.bat     dotnet publish でビルド
├── rust/         Rust 再実装（calamine 使用）
├── samples/      20,000行 × 30列 の xlsx / xls サンプル
├── bench/        COM ベンチマーク用 C# プログラム
├── bench.ps1     C# / Rust / COM のベンチマーク
└── README.md
```

- **C# 5 互換のコード**を **.NET 10 SDK** でビルド
- `dotnet build` / `dotnet format` / `dotnet test` で開発
- 出力は **NDJSON / JSON** → 絞り込みは `jq` に任せる（jq を再発明しない）
- `build.bat` は `dotnet publish` で単一ファイル `xq.exe` を生成

## ビルド（.NET 10）

```bash
dotnet build xq.csproj -c Release
dotnet format xq.csproj
dotnet test tests/xq-tests.csproj
```

Windows で単一ファイルを作る場合:

```bat
build.bat
```

`build.bat` は `dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true` を実行し、
`dist\xq.exe` をルートの `xq.exe` にコピーします。

CI では `.github/workflows/build.yml` が **build / format / test** を実行します。

## 使い方

```bat
xq sheets book.xlsx
:: ["Sheet1","Sheet2"]

xq cells book.xlsx
:: {"sheet":"Sheet1","ref":"A1","r":1,"c":1,"v":"商品名","t":"s"}
:: {"sheet":"Sheet1","ref":"B2","r":2,"c":2,"v":10,"t":"n"}
:: ...

xq cells book.xlsx --sheet Sheet1 --range A1:C10 --json
:: [ {...}, {...} ]   ← JSON 配列で欲しいとき

xq cells book.xlsx --formulas        :: 数式セルに "f" を付ける
xq cells book.xlsx --all             :: 空セルも "v":null で出す
xq cells book.xlsx --text            :: Excel の表示文字列を "text" に付ける
xq cells book.xlsx --limit 100       :: 出力するセル数を 100 件で打ち切る

xq names book.xlsx                   :: 名前定義（NDJSON）
xq names book.xlsx --name 売上       :: 名前で完全一致
xq names book.xlsx --grep 売         :: 名前 / 参照 / コメントの部分一致
xq names book.xlsx --sheet Sheet2    :: そのシートのローカル名＋ブック全体の名前
xq names book.xlsx --json            :: JSON 配列

xq cells book.xlsx --range 売上      :: 名前定義を範囲として展開
xq vba book.xlsm                     :: {"codePage":1252,"modules":[...]}
xq vba book.xlsm --module Module1    :: モジュール単体
xq info book.xlsm                    :: シート・dimension・VBA・名前定義の要約

:: バイナリ形式も同じ CLI
xq cells book.xlsb --formulas --json
xq cells book.xls  --formulas --json
xq names book.xlsb --json
xq vba   book.xls
xq info  book.xls

:: 式評価（C# 版のみ）
xq book.xlsx -e "=A1"
xq book.xlsx -e "=SUM(A1:A10)"
xq book.xlsx --json -e "=IF(A1>0,\"OK\",\"NG\")"
xq eval book.xlsx -e "=A1+B1"
```

jq との組み合わせ:

```bash
xq cells book.xlsx | jq -c 'select(.v=="りんご")'
xq cells book.xlsx | jq -r 'select(.t=="n") | "\(.ref)=\(.v)"'
xq names book.xlsx | jq -c 'select(.scope=="workbook")'
xq names book.xlsx | jq -r 'select(.name=="売上") | .ref'
xq vba book.xlsm | jq -r '.modules[] | "\(.name)\t\(.code|length)"'
xq vba book.xlsm | jq -r '.modules[] | select(.name=="Module1") | .code'
```

### セル出力のスキーマ

| フィールド | 意味 |
|---|---|
| `sheet` | シート名 |
| `ref` | A1 形式のセル番地 |
| `r` / `c` | 行 / 列（1 始まり） |
| `v` | 値。数値・真偽値は JSON の生トークン、文字列は JSON 文字列 |
| `t` | `n`(数値) `s`(文字列) `b`(真偽値) `e`(エラー) `d`(日付文字列) |
| `f` | 数式テキスト（`--formulas` 指定時のみ） |
| `text` | Excel 風の表示文字列（`--text` 指定時のみ） |

### VBA 出力のスキーマ

```json
{"codePage":1252,"modules":[
  {"name":"Module1","stream":"Module1","kind":"standard","offset":0,
   "readonly":false,"private":false,"code":"Attribute VB_Name = ..."}
]}
```

`code` は VBA のソースそのまま（CRLF、`Attribute VB_*` ヘッダ込み）。

### 名前定義の出力スキーマ（`xq names`）

```json
{"name":"売上","scope":"workbook","sheet":null,"sheetId":null,
 "hidden":false,"builtin":false,"ref":"Sheet1!$A$1:$C$5","comment":null}
```

| フィールド | 意味 |
|---|---|
| `name` | 定義名 |
| `scope` | `workbook`（ブック全体）または `sheet`（シートローカル） |
| `sheet` / `sheetId` | ローカル名の対象シート（`localSheetId`） |
| `hidden` | 非表示の名前 |
| `builtin` | `_xlnm.*`（印刷範囲・フィルタ等の組み込み名） |
| `ref` | 参照文字列（`Sheet1!$A$1:$C$5`、定数、数式など） |
| `comment` | コメント（あれば） |

### 名前定義を `--range` に使う

```bat
xq cells book.xlsx --range 売上
xq cells book.xlsx --range Sales --sheet Sheet1
```

- 解決順: `--sheet` 指定時はそのシートのローカル名 → ブック全体 → その他のローカル名
- `--sheet` を省略すると、名前が指すシートを自動で使います
- **複数領域の名前**（`Sheet1!$A$1,Sheet1!$C$1`）と**定数/数式の名前**（`Rate=0.1`）は
  単一矩形にできないためエラーにします（`xq names` では参照を確認できます）
- シート名のクォート（`'Sheet 1'!$A$1:$B$2`）と `$` は解釈します

## 表示文字列（`--text`）

`"v"` は生の値（数値・シリアル値）ですが、`--text` を付けると `styles.xml` の
`numFmt`（`cellXfs` → `numFmtId`、カスタム書式は `numFmts`）を解釈して、
Excel の `TEXT()` 相当の表示文字列を `"text"` に追加します。

```bat
xq cells book.xlsx --text
:: {"...","v":1234.5,"t":"n","text":"¥1,234.50"}
:: {"...","v":45000,"t":"n","text":"2023/3/15"}
```

対応している書式:

- `General`（11文字幅相当の丸め。`1E+11` などの指数表記を含む）
- `0` / `0.00` / `#` / `?` / `#,##0` / `#,##0.00` / `0.0#`
- 桁区切り、末尾カンマによるスケーリング（`#,##0,` = 千単位）
- `%` / `0%` / `0.00%`
- 科学技術表記（`0.00E+00`, `##0.0E+0` = 3桁区切りの工学表記）
- 通貨リテラル（`"¥"#,##0`, `[$¥-411]`, `[$€-407]`）
- 日付・時刻（`yyyy/m/d`, `yyyy-mm-dd`, `m/d/yyyy`, `yy/mm/dd`, `d-mmm-yy`,
  `mmm-yy`, `h:mm`, `h:mm:ss`, `hh:mm:ss AM/PM`, `[h]:mm:ss`, `mm:ss`,
  `yyyy"年"m"月"d"日"`, `aaa`/`aaaa` の日本語曜日）
- セクション（`正;負;ゼロ;文字`、`[Red]` などの色指定は無視）
- `.xls` / `.xlsb` も `XF` / `FORMAT`（BIFF8）や `BrtXF` / `BrtFmt`（BIFF12）から
  同じ表示文字列を出します
- 文字列（`@`, `"["@"]"`）
- 1900年うるう年バグ（シリアル 60 = 1900/2/29）と 1904 日付システム

未対応（`General` にフォールバック）:

- 分数（`# ?/?`）
- 条件付き書式（`[>100]`）
- `[DBNum1]` などの東アジア数字
- 組み込み `numFmtId` 14〜22 は en-US の並び（`mm-dd-yy` 等）。
  日本語 Excel が使う `yyyy/m/d` などは、通常はカスタム書式（164 以降）として
  ファイルに入っているので、そちらは正確に出ます。

検証: Python の `ssf`（Excel 書式エンジンの実装）と 61 ケースを比較し、
日本語曜日（`ssf` 未対応）を除いて **全一致**。`openpyxl` で作った
`datetime` セルも `2026/11/14 9:30` のように出ます。

注意: Excel の `Range.Text` は**列幅**に依存します（狭いと `####`）。
ここで出るのは `TEXT()` 相当の決定的な文字列で、列幅は考慮しません。

文字列と数値は区別します。たとえば:

```bat
:: セルに文字列 "02"（表示形式 @）
:: -> {"v":"02","t":"s","text":"02"}
:: セルに数値 2（表示形式 General）
:: -> {"v":2,"t":"n","text":"2"}
:: セルに数値 2（表示形式 00）
:: -> {"v":2,"t":"n","text":"02"}
```

`-e` はセルの**値**を返すため、セルの表示形式は適用しません。
表示形式込みの文字列が必要な場合は `xq cells --text` を使ってください。



## 式評価（`-e` / `eval`）

C# 版では、Excel を起動せずにセル参照や簡単な数式を評価できます。

```bat
xq book.xlsx -e "=A1"
xq book.xlsx -e "=A1+B1"
xq book.xlsx -e "=SUM(A1:A10)"
xq book.xlsx --json -e "=IF(A1>0,\"OK\",\"NG\")"
xq eval book.xlsx -e "=A1+B1"
```

- `xq <file> -e EXPR` と `xq eval <file> -e EXPR` は同じ動作です。
- 参照: `A1`, `$A$1`, `Sheet1!A1`, `'Sheet 1'!A1`, `A1:B2`, `F:H` / `$F:$H`（全列）, 名前定義
- 演算子: `+ - * / ^ &`、比較（`= <> < <= > >=`）、`%`、括弧
- 関数:
  - 集計: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`
  - 条件: `IF`, `IFERROR`, `AND`, `OR`, `NOT`
  - 数学: `ABS`, `INT`, `MOD`, `POWER`, `SQRT`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`
  - 文字列: `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `UPPER`, `LOWER`, `CONCATENATE`, `VALUE`, `TEXT`, `SUBSTITUTE`, `REPLACE`
  - 判定: `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISNA`
  - 検索・参照: `VLOOKUP`, `INDEX`, `MATCH`, `ADDRESS`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`
  - 条件集計: `SUMIF`, `COUNTIF`
  - 間接参照: `INDIRECT`
  - 日付: `DATE`, `YEAR`, `MONTH`, `DAY`
  - その他: `TRUE`, `FALSE`, `NA`, `PI`
- 関数は入れ子にできます。例: `=IF(A1>0,SUM(A1:A3),MAX(A1:A3))`、
  `=INDEX(A1:B3,MATCH(2,A1:A3,0),2)`、`=TEXT(ROUND(AVERAGE(A1:A3),2),"0.00")`
- 文字列 `"02"` と数値 `02` は区別します。`="02"` は文字列 `02`、`=02` は数値 `2` です。
- セルの値は、まずファイルに保存されている**キャッシュ値**を使います。
  キャッシュ値のない数式セルは、参照関係から**依存関係グラフ**を作り、
  依存先から順に（Kahn のトポロジカルソートで）評価します。循環参照は `#REF!` です。
- 参照範囲はシート単位でまとめて読み込むため、同じ範囲を何度も走査しません。
- 全列参照（`F:H`, `$F:$H`）は、そのシートの使用範囲の最終行までに丸めて評価します。
  Excel の 1,048,576 行全体をメモリに展開しません。
- `--json` を付けると `{"v":...,"t":"n|s|b|e|z|a"}` を出力します。
  `--text` を付けると `"text"` も追加します。
- エラー値: `#DIV/0!`, `#VALUE!`, `#NAME?`, `#REF!`, `#N/A`, `#NUM!`, `#NULL!`

制限:

- Excel の完全互換ではありません。
- `VLOOKUP` / `INDEX` / `MATCH` / `SUMIF` / `COUNTIF` / `INDIRECT` は実装済みですが、
  Excel の全エッジケース（ワイルドカード、近似一致、外部参照など）を完全再現するものではありません。
- 配列数式、スピル、3D参照、外部ブック参照、volatile 関数は未対応です。
- `INDIRECT` は `A1` 形式のみ対応し、`R1C1` と外部ブック参照は `#REF!` になります。
- `_xlfn.*` の新しい関数は、実装済みのもの以外は `#NAME?` になります。
- キャッシュ値がある数式セルは再計算しません（将来 `--recalc` を検討）。
- `INDIRECT` など実行時に決まる参照は依存関係グラフに静的に含めず、評価時にフォールバックします。
- Rust 版 `xq-rs` には `-e` はまだ実装していません。

## テスト

```bash
dotnet test tests/xq-tests.csproj
```

- `tests/XqTests.cs` … xUnit による評価器テスト（223ケース。参照、全列参照、演算、関数、入れ子関数、エラー、日付、依存関係グラフ、文字列/数値、`.xlsx` / `.xls` / `.xlsb`）
- `tests/run_tests.ps1` … `dotnet test` と CLI スモークテスト（`xq.exe` のビルドも行う）
- `tests/fixtures/eval.xlsx` / `eval.xls` / `eval.xlsb` … テスト用ブック
- `tests/fixtures/graph.xlsx` … 依存関係グラフ・循環参照のテスト用ブック
- `tests/fixtures/types.xlsx` … 文字列 `"02"` / 数値 `2` / 表示形式 `00` のテスト用ブック
- `tests/fixtures/fullcol.xlsx` … `シート!$F:$H` のような全列参照のテスト用ブック
- `tests/make_fixtures.py` … `eval.xlsx` / `graph.xlsx` / `types.xlsx` / `fullcol.xlsx` / `eval.xls` の再生成

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run_tests.ps1
```

## バイナリ形式（.xlsb / .xls）

- **.xlsb** … ZIP パッケージ + BIFF12（`xl/workbook.bin`, `xl/worksheets/sheetN.bin`,
  `xl/sharedStrings.bin`, `xl/styles.bin`, `xl/vbaProject.bin`）。
- **.xls** … CFB（OLE2）の `Workbook` ストリーム + BIFF8（Excel 97-2003）。
  BIFF5/95 は対象外です。
- **暗号化** … `.xls` の `FILEPASS`、および CFB の `EncryptedPackage` を検出したら
  エラーにします（復号は未実装）。
- **数式** … `Ptg` トークンを C# 5 でデコンパイルします。
  - BIFF8: `PtgRef` / `PtgArea` / `PtgRef3d` / `PtgArea3d` / `PtgName` /
    `PtgFunc` / `PtgFuncVar` / `PtgStr` / `PtgAttr` / `PtgExp`（共有数式）など。
  - BIFF12: 同じ関数表を使い、BIFF12 のペイロード長（`Ref` 7, `Area` 13,
    `Ref3d` 11, `Area3d` 17 など）で読みます。
  - 共有数式（`SHRFMLA` / `BrtShrFmla`）と配列数式（`ARRAY` / `BrtArrFmla`）を
    解決し、依存セルでは基準セルからの相対オフセットを適用します。
  - 外部シート参照は `EXTERNSHEET` / `SUPBOOK`（.xls）と
    `BrtExternSheet` / `BrtSupSelf` / `BrtBundleSh`（.xlsb）から解決します。
- **名前定義** … `NAME` / `BrtName` の `rgce` をデコンパイルして `ref` を復元します。
  範囲・複数領域・定数は xlsx と同じ形式で出ます。数式名はデコンパイラの
  対応範囲に依存します。
- **VBA** … `.xls` は CFB 全体、`.xlsb` は `xl/vbaProject.bin` を既存の
  `Cfb.cs` / `Ovba.cs` に渡します。

制限:

- `PtgNameX`（外部ブックの名前）や一部の新しい関数（`FUNC255` / `_xlfn.*`）は
  完全な Excel 表記にならないことがあります。
- 配列定数（`PtgArray`）は `{...}` のプレースホルダになります。
- `.xlsb` の `PtgExp` は列情報が `RgbExtra` にあるため、共有範囲から
  マスター列を推定しています（通常の共有数式では一致）。
- `.xls` の BIFF8 以外（BIFF5/95）と暗号化ファイルは未対応です。

## Rust 版（`rust/`）

同じ CLI を Rust で再実装したものです。読み取りには `calamine` 0.26.1 を使います。

```
rust/
├── Cargo.toml
├── src/main.rs
└── xq-rs.exe     Windows x86_64 ビルド
```

```bash
cd rust
cargo build --release
# Windows 向けクロスビルド（Alpine + mingw-w64）
cargo build --release --target x86_64-pc-windows-gnu
```

対応コマンドは `sheets` / `cells` / `names` / `vba` / `info` です。
`calamine` が `.xlsx` / `.xls` / `.xlsb` の値・数式・名前定義・VBA を読むため、
C# 版とほぼ同じ JSON を出します。ただし:

- `--text` は数値の生文字列のみ（`numFmt` の表示文字列は未実装）。
- `names` は `scope` / `hidden` / `comment` を `calamine` が公開しないため、
  `scope:"workbook"` 固定です。
- `vba` のモジュールは `name` と `code` のみ（`kind` / `offset` なし）。
- `calamine` の `.xls` 数式デコンパイラは一部の 3D 参照を
  `$USM$` のような誤った列名にすることがあります。

## サンプルとベンチマーク

20,000 行 × 30 列のサンプルを `samples/` に置いています。

```
samples/sample_20000x30.xlsx   (約 2.3 MB)
samples/sample_20000x30.xls    (約 8.4 MB)
samples/make_samples.py        再生成スクリプト
```

ベンチマークは `bench.ps1` です。C# 版・Rust 版・COM 版
（`bench/ComRead.cs` と、Excel COM が使えない場合の ADO/ACE
フォールバック `bench/ComReadAdo.cs`）を同じファイルに対して実行し、
`bench-results.csv` を出します。

```powershell
powershell -ExecutionPolicy Bypass -File .\bench.ps1 -Runs 3
```

- C# / Rust は `xq cells <file> --json > NUL` の wall time。
- COM はまず Excel COM（`Excel.Application` の `UsedRange.Value2`）を試し、
  失敗した場合は ADO/ACE COM（`ADODB.Connection` +
  `Microsoft.ACE.OLEDB.16.0` の `GetRows`）にフォールバックします。
- この検証 VM の Excel はインストール済みですが未アクティベーションで
  `Workbooks.Open` が失敗するため、`.xlsx` は ADO/ACE COM で測定しました。
  `.xls` は `xlwt` 生成ファイルを ACE が読めないため `skipped` です。
  Excel をアクティベーションした環境では Excel COM で `.xls` も測れます。

実測（Windows 11 VM, 8 vCPU / 10 GB, .NET Framework 4.8）:

| ファイル | C# `xq.exe` | Rust `xq-rs.exe` | COM |
|---|---:|---:|---:|
| `sample_20000x30.xlsx` | **約 640 ms** | 約 1,090 ms | 約 1,593 ms（ADO/ACE COM） |
| `sample_20000x30.xls`  | **約 515 ms** | 約 868 ms  | skipped（ACE が xlwt `.xls` を読めない） |

C# 版は JSON シリアライズ込みで Rust 版より約 1.7 倍高速でした。
COM 版を測る場合は、Excel をアクティベーションした Windows で
`bench.ps1` を実行してください。

`--limit 100` の実測（同じサンプル）:

| 実装 | 時間 |
|---|---:|
| C# `xq.exe` | **約 68 ms** |
| Rust `xq-rs.exe` | 約 372 ms |

C# 版は 100 件出力した時点で読み取りを打ち切ります。
Rust 版は `calamine` がシート全体をパースするため、出力だけを 100 件に制限します。

## 実装（なぜ速いか）

- **COM を使わない**。Excel のプロセス間呼び出し（1セル数 ms）を排除。
- **ZIP をストリームで読む**（`System.IO.Compression.ZipArchive`）。
- **セルは `XmlReader` でストリーミング**。DOM を作らず、行単位で捨てる。
- **共有文字列は配列に展開**して O(1) 参照。
- **範囲指定はプッシュダウン**（`--range` で行・列を先に判定し、範囲外は `Skip()`）。
- **NDJSON を 64KB バッファに溜めて書き出し**。
- **VBA は自前実装**:
  - `Cfb.cs` … MS-CFB（OLE2）リーダ。FAT / mini-FAT / ディレクトリ / mini ストリーム。
  - `Ovba.cs` … MS-OVBA の展開（raw チャンク・コピートークン・オーバーラップコピー）と
    `dir` ストリームの解析（モジュール名・ストリーム名・`MODULEOFFSET`・コードページ）。
  - アルゴリズムは **pyOpenVBA (MIT)** をリバースエンジニアリングして C# 5 に移植。
- **セル読み**は openpyxl の OOXML 解釈（sharedStrings / inlineStr / `t` 属性 / 数式）に準拠。

## 実測（Windows 11 VM, 8 vCPU / 10 GB, .NET Framework 4.8）

| 処理 | xq | 比較 |
|---|---|---|
| 起動（`--version`） | **約 35 ms** | — |
| 300,000 セル（50k行×6列）NDJSON | **約 315 ms**（≒95万セル/秒） | openpyxl 2,697 ms（**約8.6倍速**） |
| 同上 + `--text` | **約 415 ms** | 書式解釈の追加コストは約 60 ms |
| 6,000 セル（`--range A1:F1000`） | **約 66 ms** | — |
| VBA 5モジュール（25 KB、複数チャンク） | **約 65 ms** | pyOpenVBA 2 ms（プロセス内）＋Python起動 |
| COM で同規模 | 数分〜（1セル数ms） | 桁違いに遅い |

正しさの検証:
- 5,000 セルを openpyxl とランダム比較 → **不一致 0**
- VBA 5 モジュールを pyOpenVBA と比較 → **全モジュール完全一致**（25,660 文字の複数チャンク含む）

## 制限・注意

- **読み取り専用**。セル/VBA の書き込みは未実装（次フェーズ）。
- **.xlsx / .xlsm のみ**。`.xlsb`（バイナリ）と `.xls`（BIFF8）は未対応。
- **日付はシリアル値のまま**（`numFmt` / スタイル解決は未実装）。
- **PowerShell の `>` は UTF-16 に変換します**。ファイルへ落とすときは:

  ```bat
  cmd /c "xq cells book.xlsx > out.ndjson"
  ```

  または `xq ... | Out-File -Encoding utf8 out.ndjson`。
- 起動 35 ms は .NET Framework の CLR 初期化由来。`ngen install xq.exe` を試しても
  ほとんど変わりませんでした（JIT より CLR 起動が支配的）。さらなる短縮は
  Native AOT / Rust / Go が必要。
- ソースは C# 5 の範囲なので、**同じコードを .NET 10 でビルドすれば Linux でも動きます**
  （その場合は Native AOT で起動 5 ms 前後も可能）。

## ファイル

| ファイル | 役割 |
|---|---|
| `src/xq.cs` | CLI・JSON 出力 |
| `src/NumFmt.cs` | Excel 表示文字列エンジン（`--text`） |
| `src/Xlsx.cs` | ZIP / workbook / sharedStrings / セル走査 / 名前定義 |
| `src/Cfb.cs` | MS-CFB リーダ |
| `src/Ovba.cs` | MS-OVBA 展開 + `dir` 解析 |
| `src/A1.cs` | A1 番地・範囲の解析 |
| `build.bat` | C# 5 ビルド |
| `xq.exe` | ビルド済み（44 KB） |

## 次のフェーズ候補

1. **書き込み**（セル: sheet XML の最小パッチ / VBA: MS-OVBA 圧縮 + `dir` 再構築）
2. `.xlsb` / `.xls` 対応
3. 分数・条件付き書式・ロケール別組み込み書式
4. `--where` などの簡易フィルタ（それでも jq 連携が主）
5. Linux / Native AOT ビルド（同じソースで）
