param(
    [string[]]$Files = @("samples\sample_20000x30.xlsx", "samples\sample_20000x30.xls"),
    [int]$Runs = 3,
    [string]$OutCsv = "bench-results.csv"
)

$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot

# Build the COM benchmark helper once.
$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (!(Test-Path $csc)) { $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe" }
if ((Test-Path "bench\ComRead.cs") -and !(Test-Path "bench\ComRead.exe")) {
    & $csc /nologo /optimize+ /target:exe /out:bench\ComRead.exe bench\ComRead.cs
}
if ((Test-Path "bench\ComReadAdo.cs") -and !(Test-Path "bench\ComReadAdo.exe")) {
    & $csc /nologo /optimize+ /target:exe /out:bench\ComReadAdo.exe bench\ComReadAdo.cs
}

$rows = New-Object System.Collections.Generic.List[object]

function Add-Result([string]$file, [string]$tool, [int]$run, [double]$ms, [string]$note) {
    $rows.Add([pscustomobject]@{
        file = $file
        tool = $tool
        run  = $run
        ms   = [math]::Round($ms, 1)
        note = $note
    })
}

foreach ($file in $Files) {
    if (!(Test-Path $file)) {
        Write-Warning "missing file: $file"
        continue
    }

    Write-Host "=== $file ==="

    $tools = @(
        @{ Name = "C#";   Exe = ".\xq.exe";    Args = "cells `"$file`" --json" },
        @{ Name = "Rust"; Exe = ".\xq-rs.exe"; Args = "cells `"$file`" --json" }
    )

    foreach ($tool in $tools) {
        # Warm-up (file cache, JIT/loader).
        cmd /c "$($tool.Exe) $($tool.Args) > NUL" | Out-Null

        $times = @()
        for ($i = 1; $i -le $Runs; $i++) {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            cmd /c "$($tool.Exe) $($tool.Args) > NUL" | Out-Null
            $sw.Stop()
            $ms = $sw.Elapsed.TotalMilliseconds
            $times += $ms
            Add-Result $file $tool.Name $i $ms ""
        }
        $min = ($times | Measure-Object -Minimum).Minimum
        $avg = ($times | Measure-Object -Average).Average
        Write-Host ("{0,-5} min={1,8:N1} ms  avg={2,8:N1} ms" -f $tool.Name, $min, $avg)
    }

    # Prefer Excel COM; fall back to ADO/ACE COM (both are COM readers).
    $comExe = $null
    $comLabel = $null
    if (Test-Path "bench\ComRead.exe") {
        $probe = & .\bench\ComRead.exe $file 2>&1
        if ($LASTEXITCODE -eq 0) { $comExe = "bench\ComRead.exe"; $comLabel = "Excel COM" }
    }
    if (!$comExe -and (Test-Path "bench\ComReadAdo.exe")) {
        $probe = & .\bench\ComReadAdo.exe $file 2>&1
        if ($LASTEXITCODE -eq 0) { $comExe = "bench\ComReadAdo.exe"; $comLabel = "ADO/ACE COM" }
    }

    if ($comExe) {
        $comOk = $true
        for ($i = 1; $i -le $Runs; $i++) {
            $out = & ".\$comExe" $file 2>&1
            if ($LASTEXITCODE -ne 0) {
                $comOk = $false
                Add-Result $file "COM" $i 0 ("skipped: " + ($out -join " "))
                break
            }
            $obj = $out | ConvertFrom-Json
            Add-Result $file "COM" $i ([double]$obj.ms) ("engine=" + $comLabel + "; cells=" + $obj.count)
            Write-Host ("COM   run {0}: {1} ms ({2}, cells={3})" -f $i, $obj.ms, $comLabel, $obj.count)
        }
        if (!$comOk) { Write-Host "COM   skipped (COM reader failed)" }
    } else {
        Add-Result $file "COM" 0 0 "skipped: Excel COM and ADO/ACE COM unavailable"
        Write-Host "COM   skipped (Excel COM and ADO/ACE COM unavailable)"
    }
}

$rows | Export-Csv -NoTypeInformation -Encoding UTF8 $OutCsv
Write-Host "wrote $OutCsv"
$rows | Format-Table -AutoSize
