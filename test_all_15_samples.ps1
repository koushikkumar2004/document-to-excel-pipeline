$ErrorActionPreference = "Continue"

$results = @()
$start = Get-Date

for ($i = 1; $i -le 15; $i++) {
    $file = if ($i -eq 1 -and (Test-Path "test_files\sample_invoice.pdf")) {
        "test_files\sample_invoice.pdf"
    } else {
        "test_files\sample_invoice$i.pdf"
    }

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $excel = "test_files\$stem.xlsx"

    Write-Host ""
    Write-Host "============================================================"
    Write-Host "TESTING SAMPLE $i OF 15 : $file"
    Write-Host "============================================================"

    if (-not (Test-Path $file)) {
        Write-Host "MISSING FILE: $file" -ForegroundColor Yellow
        $results += [PSCustomObject]@{
            Sample = $i
            Filename = [System.IO.Path]::GetFileName($file)
            Status = "MISSING"
            Seconds = 0
            Excel = ""
        }
        continue
    }

    # Remove stale Excel output before test to avoid false positive classification
    if (Test-Path $excel) {
        Remove-Item $excel -Force -ErrorAction SilentlyContinue
    }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    python pipeline.py $file
    $exitCode = $LASTEXITCODE
    $timer.Stop()

    $status = if ($exitCode -eq 0 -and (Test-Path $excel)) { "PROCESSED" } else { "FAILED" }

    $results += [PSCustomObject]@{
        Sample = $i
        Filename = [System.IO.Path]::GetFileName($file)
        Status = $status
        Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 2)
        Excel = if (Test-Path $excel) { $excel } else { "" }
    }
}

Write-Host ""
Write-Host "================ FINAL TEST SUMMARY ================"
$results | Format-Table -AutoSize
$summaryPath = "test_files\milestone7_test_summary.csv"
$results | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
Write-Host "Summary saved to: $summaryPath"
Write-Host ("Total elapsed: {0:N2} minutes" -f ((Get-Date) - $start).TotalMinutes)
