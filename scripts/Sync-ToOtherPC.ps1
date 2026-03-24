#Requires -Version 5.1
<#
.SYNOPSIS
  매니페스트에 적힌 파일만 다른 PC 프로젝트 루트로 복사합니다.
.EXAMPLE
  .\Sync-ToOtherPC.ps1 -Destination "D:\작업\방과후센터앱"
  .\Sync-ToOtherPC.ps1 -Destination "D:\작업\방과후센터앱" -Mode App
#>
param(
    [Parameter(Mandatory = $true)]
    [string] $Destination,
    [ValidateSet("Changes", "App")]
    [string] $Mode = "Changes"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = (Resolve-Path (Join-Path $ScriptDir "..")).Path
$ManifestName = if ($Mode -eq "App") { "sync-manifest-app전체.txt" } else { "sync-manifest-변경분.txt" }
$ManifestPath = Join-Path $ScriptDir $ManifestName

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    Write-Error "매니페스트 없음: $ManifestPath"
}

$DestRoot = $Destination.TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $DestRoot)) {
    New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null
}

$lines = Get-Content -LiteralPath $ManifestPath -Encoding UTF8
$n = 0
foreach ($line in $lines) {
    $rel = $line.Trim()
    if ($rel -eq "" -or $rel.StartsWith("#")) { continue }
    $src = Join-Path $ProjectRoot $rel
    $dst = Join-Path $DestRoot $rel
    if (-not (Test-Path -LiteralPath $src)) {
        Write-Warning "건너뜀: $rel"
        continue
    }
    $dstDir = Split-Path -Parent $dst
    if (-not (Test-Path -LiteralPath $dstDir)) {
        New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $src -Destination $dst -Force
    $n++
    Write-Host "복사: $rel"
}
Write-Host "완료 $n 개 -> $DestRoot"
