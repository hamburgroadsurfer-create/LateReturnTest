$ErrorActionPreference = "Stop"

$htmlPath = "C:\Users\AlexanderOchsRoadsur\Roadsurfer GmbH\RSF - RSF\46_Fleet Security\Fleet Monitor Test\ReturnReportTest.html"
$jsPath = "C:\Users\AlexanderOchsRoadsur\Roadsurfer GmbH\RSF - RSF\46_Fleet Security\Fleet Monitor Test\report_data.js"
$tcuPath = "C:\Users\AlexanderOchsRoadsur\Roadsurfer GmbH\RSF - Connected Vehicle\FleetData\Exports\TCU"
$appPath = "C:\Users\AlexanderOchsRoadsur\Roadsurfer GmbH\RSF - RSF\46_Fleet Security\Fleet Monitor Test"

Write-Host "🔍 Searching for latest TCU file..." -ForegroundColor Cyan
$latestTcu = Get-ChildItem "$tcuPath\TCU health vehicle data_*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $latestTcu) {
    Write-Host "❌ Error: No TCU file found in $tcuPath" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "✅ Found TCU: $($latestTcu.Name)" -ForegroundColor Green

function Get-Base64($path) {
    if (Test-Path $path) {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        return [Convert]::ToBase64String($bytes)
    }
    Write-Host "⚠️ Warning: File not found: $path" -ForegroundColor Yellow
    return $null
}

Write-Host "📦 Packaging files..." -ForegroundColor Cyan

$tcuB64 = Get-Base64 $latestTcu.FullName
$stationsB64 = Get-Base64 "$appPath\Stations_Merged.csv"
$bookingsB64 = Get-Base64 "$appPath\Bookings_Return_Today.csv"
$carsB64 = Get-Base64 "$appPath\Cars without Booking with millage diff.csv"

# Construct JS Content
$jsContent = @"
window.reportData = {
  tcu: { name: "$($latestTcu.Name)", data: "$tcuB64" },
  stations: { name: "Stations_Merged.csv", data: "$stationsB64" },
  bookings: { name: "Bookings_Return_Today.csv", data: "$bookingsB64" },
  cars: { name: "Cars_without_Booking.csv", data: "$carsB64" }
};
"@

$jsContent | Out-File -FilePath $jsPath -Encoding utf8
Write-Host "✅ Data injected into report_data.js" -ForegroundColor Green

Write-Host "🚀 Launching Report..." -ForegroundColor Cyan
Start-Process $htmlPath
