$ErrorActionPreference = "Continue"

$root = Split-Path $PSScriptRoot -Parent
Set-Location $root

$failures = 0

# xUnit tests (primary test suite).
dotnet test tests\xq-tests.csproj -c Release
if ($LASTEXITCODE -ne 0) {
    $failures++
    Write-Host "dotnet test FAILED"
} else {
    Write-Host "dotnet test PASSED"
}

# CLI smoke tests need the published xq.exe.
if (!(Test-Path (Join-Path $root "xq.exe"))) {
    cmd /c build.bat
    if ($LASTEXITCODE -ne 0) {
        Write-Host "build FAILED"
        exit 1
    }
}

function Invoke-Xq([string]$rawArgs) {
    $outFile = Join-Path $PSScriptRoot "out.txt"
    $errFile = Join-Path $PSScriptRoot "err.txt"
    Remove-Item $outFile, $errFile -ErrorAction SilentlyContinue

    $p = Start-Process -FilePath (Join-Path $root "xq.exe") `
        -ArgumentList $rawArgs -Wait -NoNewWindow -PassThru `
        -RedirectStandardOutput $outFile -RedirectStandardError $errFile

    $out = ""
    if (Test-Path $outFile) {
        $tmp = Get-Content $outFile -Raw -ErrorAction SilentlyContinue
        if ($null -ne $tmp) { $out = $tmp.Trim() }
    }
    $err = ""
    if (Test-Path $errFile) {
        $tmp = Get-Content $errFile -Raw -ErrorAction SilentlyContinue
        if ($null -ne $tmp) { $err = $tmp.Trim() }
    }
    $code = -1
    if ($null -ne $p) { $code = $p.ExitCode }
    return @{ Code = $code; Out = $out; Err = $err }
}

function Check-Cli([string]$file, [string]$expr, [string]$expected) {
    $escaped = $expr.Replace('"', '\"')
    $r = Invoke-Xq ($file + ' -e "' + $escaped + '"')
    if ($r.Code -ne 0 -or $r.Out -ne $expected) {
        $script:failures++
        Write-Host ("CLI FAIL: {0} -e {1} -> [{2}] code={3} expected [{4}]" -f $file, $expr, $r.Out, $r.Code, $expected)
    } else {
        Write-Host ("CLI OK:   {0} -e {1} -> {2}" -f $file, $expr, $r.Out)
    }
}

$cliCases = @(
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=A1"; Expected = "1" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=A1+A2"; Expected = "3" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=SUM(A1:A3)"; Expected = "6" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=Other!A1"; Expected = "10" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=IF(A1>0,"yes","no")'; Expected = "yes" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=TEXT(1234.5,"#,##0.00")'; Expected = "1,234.50" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=CONCATENATE(B1," ",B2)'; Expected = "Hello World" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=VLOOKUP(2,A1:B3,2,FALSE)"; Expected = "World" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=INDEX(A1:B3,2,2)"; Expected = "World" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=MATCH(2.5,A1:A3,1)"; Expected = "2" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=SUMIF(A1:A3,">1",K1:K3)'; Expected = "50" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=COUNTIF(B1:B3,"*o*")'; Expected = "2" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=INDIRECT("A1:A3")'; Expected = "1" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=ADDRESS(2,3,4)"; Expected = "C2" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=ROW(A5:B7)"; Expected = "5" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=COLUMN(C3)"; Expected = "3" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=ROWS(A1:B3)"; Expected = "3" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=COLUMNS(A1:B3)"; Expected = "2" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=SUBSTITUTE("a-b-c","-","+",2)'; Expected = "a-b+c" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = '=REPLACE("abcdef",2,3,"X")'; Expected = "aXef" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=C1"; Expected = "3" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=C2"; Expected = "6" },
    @{ File = "tests\fixtures\graph.xlsx"; Expr = "=A3"; Expected = "4" },
    @{ File = "tests\fixtures\graph.xlsx"; Expr = "=B2"; Expected = "11" },
    @{ File = "tests\fixtures\graph.xlsx"; Expr = "=Other!A1"; Expected = "11" },
    @{ File = "tests\fixtures\graph.xlsx"; Expr = "=C1"; Expected = "#REF!" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = '="02"'; Expected = "02" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = "=02"; Expected = "2" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = '=TEXT(2,"00")'; Expected = "02" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = '=VALUE("02")'; Expected = "2" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = "=A1"; Expected = "02" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = "=A2"; Expected = "2" },
    @{ File = "tests\fixtures\types.xlsx"; Expr = "=A3"; Expected = "2" },
    @{ File = "tests\fixtures\eval.xlsx"; Expr = "=1/0"; Expected = "#DIV/0!" },
    @{ File = "tests\fixtures\eval.xls"; Expr = "=A1+A2"; Expected = "3" },
    @{ File = "tests\fixtures\eval.xlsb"; Expr = "=SUM(col1_)"; Expected = "15" }
)

foreach ($c in $cliCases) {
    Check-Cli $c.File $c.Expr $c.Expected
}

$r = Invoke-Xq 'tests\fixtures\eval.xlsx --json -e "=A1"'
if ($r.Code -eq 0) {
    try {
        $obj = $r.Out | ConvertFrom-Json
        if ([double]$obj.v -ne 1 -or $obj.t -ne "n") {
            $failures++
            Write-Host ("JSON FAIL: {0}" -f $r.Out)
        } else {
            Write-Host ("JSON OK:   {0}" -f $r.Out)
        }
    } catch {
        $failures++
        Write-Host ("JSON FAIL: {0}" -f $r.Out)
    }
} else {
    $failures++
    Write-Host ("JSON FAIL: exit={0} err={1}" -f $r.Code, $r.Err)
}

$r = Invoke-Xq 'cells tests\fixtures\types.xlsx --text --json'
if ($r.Code -eq 0) {
    try {
        $arr = $r.Out | ConvertFrom-Json
        $a1 = ($arr | Where-Object { $_.ref -eq "A1" }).text
        $a2 = ($arr | Where-Object { $_.ref -eq "A2" }).text
        $a3 = ($arr | Where-Object { $_.ref -eq "A3" }).text
        if ($a1 -ne "02" -or $a2 -ne "2" -or $a3 -ne "02") {
            $failures++
            Write-Host ("TEXT FAIL: A1=[{0}] A2=[{1}] A3=[{2}]" -f $a1, $a2, $a3)
        } else {
            Write-Host ("TEXT OK:   A1={0} A2={1} A3={2}" -f $a1, $a2, $a3)
        }
    } catch {
        $failures++
        Write-Host ("TEXT FAIL: {0}" -f $r.Out)
    }
} else {
    $failures++
    Write-Host ("TEXT FAIL: exit={0} err={1}" -f $r.Code, $r.Err)
}

$r = Invoke-Xq 'tests\fixtures\eval.xlsx -e "=A1+"'
if ($r.Code -ne 2) {
    $failures++
    Write-Host ("PARSE FAIL: exit={0} expected 2" -f $r.Code)
} else {
    Write-Host "PARSE OK: exit=2"
}

$r = Invoke-Xq 'tests\fixtures\eval.xlsx --sheet Other -e "=A1"'
if ($r.Code -ne 0 -or $r.Out -ne "10") {
    $failures++
    Write-Host ("SHEET FAIL: [{0}] code={1}" -f $r.Out, $r.Code)
} else {
    Write-Host "SHEET OK: 10"
}

Write-Host ""
Write-Host ("failures=" + $failures)
if ($failures -ne 0) { exit 1 }
exit 0
