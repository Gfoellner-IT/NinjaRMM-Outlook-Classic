<#
.SYNOPSIS
Builds the M9 solution and tests the core/infrastructure. Never starts Office.
.DESCRIPTION
Default: Visual Studio MSBuild and .NET Framework 4.8 reference assemblies are required.
-RuntimeOnly: limited syntax/behavior probe using the installed Framework compiler and runtime
assemblies. This does NOT validate the net48 project, target reference API or VSTO host.
No installation, package restore, certificate, security or credential changes.
#>
[CmdletBinding()]
param([switch]$RuntimeOnly)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$output = Join-Path $repoRoot 'artifacts/core-runtime'
if ($RuntimeOnly) {
    Write-Output 'DIAGNOSTIC ONLY: installed runtime references, not a net48/VSTO build.'
    $compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    if (-not (Test-Path -LiteralPath $compiler)) { throw 'Framework C# compiler not found.' }
    $null = New-Item -ItemType Directory -Path $output -Force
    $sources = @(Get-ChildItem (Join-Path $repoRoot 'dotnet/NinjaRmm.Outlook.Core') -Filter '*.cs' -File | Sort-Object Name | ForEach-Object FullName)
    $library = Join-Path $output 'NinjaRmm.Outlook.Core.dll'
    $tests = Join-Path $output 'NinjaRmm.Outlook.Core.Tests.exe'
    & $compiler /nologo /warn:4 /warnaserror+ /reference:System.Web.Extensions.dll /target:library "/out:$library" @sources
    if ($LASTEXITCODE) { throw 'Core diagnostic compilation failed.' }
    $testSources = @(Get-ChildItem (Join-Path $repoRoot 'dotnet/NinjaRmm.Outlook.Core.Tests') -Filter '*.cs' -File | Sort-Object Name | ForEach-Object FullName)
    & $compiler /nologo /warn:4 /warnaserror+ /target:exe "/reference:$library" "/out:$tests" @testSources
    if ($LASTEXITCODE) { throw 'Test diagnostic compilation failed.' }
} else {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $referenceDir = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8'
    if (-not (Test-Path -LiteralPath $vswhere) -or -not (Test-Path -LiteralPath (Join-Path $referenceDir 'mscorlib.dll'))) {
        throw 'Visual Studio MSBuild and .NET Framework 4.8 targeting pack required. See docs/m9/architecture.md. RuntimeOnly is diagnostic only.'
    }
    $msbuild = @(& $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe') | Select-Object -First 1
    if (-not $msbuild) { throw 'Visual Studio MSBuild not found.' }
    & $msbuild (Join-Path $repoRoot 'dotnet/NinjaRmm.Outlook.sln') /nologo /t:Rebuild /p:Configuration=Release /p:Platform='Any CPU' /v:minimal
    if ($LASTEXITCODE) { throw 'net48 build failed.' }
    $tests = Join-Path $repoRoot 'dotnet/NinjaRmm.Outlook.Core.Tests/bin/Release/NinjaRmm.Outlook.Core.Tests.exe'
}
$fixtureOutput = Join-Path (Split-Path $tests) 'fixtures'
$null = New-Item -ItemType Directory -Path $fixtureOutput -Force
Copy-Item -Path (Join-Path $repoRoot 'Tests/fixtures/*.json') -Destination $fixtureOutput
& $tests
if ($LASTEXITCODE) { throw 'Core reference tests failed.' }

if (-not $RuntimeOnly) {
    $infrastructureTests = Join-Path $repoRoot 'dotnet/NinjaRmm.Outlook.Infrastructure.Tests/bin/Release/NinjaRmm.Outlook.Infrastructure.Tests.exe'
    & $infrastructureTests
    if ($LASTEXITCODE) { throw 'Infrastructure reference tests failed.' }
}
