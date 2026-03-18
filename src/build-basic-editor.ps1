param(
    [string]$OutputName = "SharpMZBasicProgramEditor.exe",
    [string]$IconPath = "c:\Users\ricky\Downloads\mz-ship.ico"
)

$ErrorActionPreference = "Stop"

$compiler = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $compiler)) {
    throw "C# compiler not found at $compiler"
}

$frameworkRoot = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
$wpfRoot = Join-Path $frameworkRoot "WPF"

$references = @(
    (Join-Path $wpfRoot "PresentationCore.dll"),
    (Join-Path $wpfRoot "PresentationFramework.dll"),
    (Join-Path $wpfRoot "WindowsBase.dll"),
    (Join-Path $frameworkRoot "System.Xaml.dll"),
    (Join-Path $frameworkRoot "System.dll"),
    (Join-Path $frameworkRoot "System.Core.dll")
) | ForEach-Object { "/r:$_" }

$projectRoot = Split-Path -Parent $PSScriptRoot
$source = Join-Path $PSScriptRoot "SharpMZBasicProgramEditor.cs"
$font = Join-Path $PSScriptRoot "SharpMZ.ttf"
$output = Join-Path $projectRoot $OutputName
$iconArg = @()

if (Test-Path $IconPath) {
    $iconArg = "/win32icon:$IconPath"
}

if (-not (Test-Path $font)) {
    throw "Embedded font source not found at $font"
}

& $compiler `
    /nologo `
    /target:winexe `
    /platform:anycpu `
    /optimize+ `
    "/out:$output" `
    $iconArg `
    "/resource:$font,SharpMZ.ttf" `
    $references `
    $source

if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

Write-Host "Built $output"
