# PowerShell 7で実行：PowerShell 5 を起動してスクリプトを実行

$scriptPath = Join-Path $PSScriptRoot "Export-AccessTableDefinitions.ps1"

# PowerShell 5 のパス（環境によって異なる場合は修正）
$ps5 = "${env:WINDIR}\System32\WindowsPowerShell\v1.0\powershell.exe"

if (!(Test-Path $ps5)) {
    Write-Error "PowerShell 5 が見つかりません： $ps5"
    exit 1
}

Write-Host "▶ PowerShell 5 経由で Access スキーマを抽出中..."
Start-Process -FilePath $ps5 -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-File", "`"$scriptPath`"" -Wait
Write-Host "✅ 完了しました。"
