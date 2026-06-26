Set-Location "D:\claude\ats-web"
$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = "python"
$pinfo.Arguments = "-m uvicorn app:app --host 0.0.0.0 --port 8765"
$pinfo.WorkingDirectory = "D:\claude\ats-web"
$pinfo.UseShellExecute = $false
$proc = [System.Diagnostics.Process]::Start($pinfo)
Start-Sleep -Seconds 3
if (-not $proc.HasExited) {
    Write-Host "ATS Scorer running at http://localhost:8765" -ForegroundColor Green
    Start-Process "http://localhost:8765"
    $proc.WaitForExit()
} else {
    Write-Host "Server failed to start." -ForegroundColor Red
}
