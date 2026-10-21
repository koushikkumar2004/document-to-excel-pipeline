$ErrorActionPreference = "Continue"

$results = @()
$start = Get-Date

for ($i = 1; $i -le 15; $i++) {
    $file = "test_files\sample_invoice$i.pdf"
    Write-Host ""
    Write-Host "============================================================"
    Write-Host "TESTING SAMPLE $i OF 15 : $file"
    Write-Host "============================================================"

    if (-not (Test-Path $file)) {
        Write-Host "MISSING FILE: $file" -ForegroundColor Yellow
        $results += [PSCustomObject]@{
            Sample = $i
            Status = "MISSING"
            Seconds = 0
            Excel = ""
        }
        continue
    }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    python pipeline.py $file
    $exitCode = $LASTEXITCODE
    $timer.Stop()

    $excel = "test_files\sample_invoice$i.xlsx"
    $status = if ($exitCode -eq 0 -and (Test-Path $excel)) { "PROCESSED" } else { "FAILED" }

    $results += [PSCustomObject]@{
        Sample = $i
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
