#Requires -Version 5.1

# Install: open PowerShell as Administrator in this folder, then run:
# powershell.exe -ExecutionPolicy Bypass -File ".\install.ps1"
# Requires WinGet and an internet connection; registers the spietopdf: URL protocol.
# Adds both PDF tools to the system PATH; reopen terminals after installation.

$ErrorActionPreference = 'Stop'

$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this installer from an elevated PowerShell window (Run as administrator).'
}

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$tempPdfDirectory = 'C:\temp_pdf'
New-Item -ItemType Directory -Path $tempPdfDirectory -Force | Out-Null
$tempPdfAcl = Get-Acl -LiteralPath $tempPdfDirectory
$tempPdfRule = [Security.AccessControl.FileSystemAccessRule]::new(
    $principal.Identity.User,
    [Security.AccessControl.FileSystemRights]::Modify,
    [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
    [Security.AccessControl.PropagationFlags]::None,
    [Security.AccessControl.AccessControlType]::Allow
)
$tempPdfAcl.SetAccessRule($tempPdfRule)
Set-Acl -LiteralPath $tempPdfDirectory -AclObject $tempPdfAcl

if (-not (Test-Path (Join-Path $scriptDirectory 'pagecount.ps'))) {
    throw "Required file pagecount.ps was not found in $scriptDirectory."
}
$registryFile = Join-Path $scriptDirectory 'printerRegister.reg'
if (-not (Test-Path $registryFile)) {
    throw "Required file printerRegister.reg was not found in $scriptDirectory."
}
$batchFile = Join-Path $scriptDirectory 'pdf.bat'
if (-not (Test-Path $batchFile)) {
    throw "Required file pdf.bat was not found in $scriptDirectory."
}

$winget = Get-Command winget.exe -ErrorAction SilentlyContinue
if (-not $winget) {
    throw 'WinGet is required. Install or update App Installer from the Microsoft Store, then run this script again.'
}

$wkhtmltopdfPath = Join-Path $env:ProgramFiles 'wkhtmltopdf\bin\wkhtmltopdf.exe'
if (-not (Test-Path $wkhtmltopdfPath)) {
    Write-Host 'Installing wkhtmltopdf...'
    & $winget.Source install --id wkhtmltopdf.wkhtmltox --exact --scope machine --silent --accept-source-agreements --accept-package-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "WinGet failed to install wkhtmltopdf (exit code $LASTEXITCODE)."
    }
}

$ghostscriptRoot = Join-Path $env:ProgramFiles 'gs'
$ghostscriptPath = Get-ChildItem -Path (Join-Path $ghostscriptRoot '*\bin\gswin64c.exe') -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName
if (-not $ghostscriptPath) {
    Write-Host 'Downloading the latest official Ghostscript installer...'
    $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/ArtifexSoftware/ghostpdl-downloads/releases/latest' -Headers @{ 'User-Agent' = 'report-to-pdf-installer' }
    $installerAsset = $release.assets |
        Where-Object { $_.name -match '^gs\d+w64\.exe$' } |
        Select-Object -First 1
    if (-not $installerAsset) {
        throw 'The latest official Ghostscript release did not include a Windows 64-bit installer.'
    }

    $expectedHash = [regex]::Match($installerAsset.digest, '^sha256:([0-9a-fA-F]{64})$').Groups[1].Value
    if (-not $expectedHash) {
        throw 'The Ghostscript installer release did not provide a SHA-256 digest; refusing to install an unverifiable download.'
    }

    $installerPath = Join-Path $env:TEMP $installerAsset.name
    try {
        Invoke-WebRequest -Uri $installerAsset.browser_download_url -OutFile $installerPath
        $actualHash = (Get-FileHash -LiteralPath $installerPath -Algorithm SHA256).Hash
        if ($actualHash -ne $expectedHash) {
            throw 'The downloaded Ghostscript installer failed its SHA-256 verification.'
        }

        Write-Host 'Installing Ghostscript silently...'
        $installerProcess = Start-Process -FilePath $installerPath -ArgumentList '/S' -Wait -PassThru
        if ($installerProcess.ExitCode -ne 0) {
            throw "Ghostscript setup failed (exit code $($installerProcess.ExitCode))."
        }
    }
    finally {
        Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path $wkhtmltopdfPath)) {
    throw "wkhtmltopdf was not found at the expected location: $wkhtmltopdfPath"
}
$ghostscriptPath = Get-ChildItem -Path (Join-Path $ghostscriptRoot '*\bin\gswin64c.exe') -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 -ExpandProperty FullName
if (-not $ghostscriptPath) {
    throw "Ghostscript was not found under $ghostscriptRoot."
}

$toolDirectories = @(
    (Split-Path -Parent $wkhtmltopdfPath)
    (Split-Path -Parent $ghostscriptPath)
) | Select-Object -Unique
$machinePathEntries = @([Environment]::GetEnvironmentVariable('Path', 'Machine') -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$otherMachinePathEntries = @($machinePathEntries | Where-Object {
    $entryPath = [Environment]::ExpandEnvironmentVariables($_.Trim().Trim('"')).TrimEnd('\')
    $toolDirectories -notcontains $entryPath
})
[Environment]::SetEnvironmentVariable('Path', (@($toolDirectories) + $otherMachinePathEntries) -join ';', 'Machine')

$currentPathEntries = @($env:Path -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$otherCurrentPathEntries = @($currentPathEntries | Where-Object {
    $entryPath = [Environment]::ExpandEnvironmentVariables($_.Trim().Trim('"')).TrimEnd('\')
    $toolDirectories -notcontains $entryPath
})
$env:Path = (@($toolDirectories) + $otherCurrentPathEntries) -join ';'

$wkhtmltopdfCommand = Get-Command wkhtmltopdf.exe -ErrorAction SilentlyContinue
if (-not $wkhtmltopdfCommand) {
    throw 'wkhtmltopdf is not resolvable from PATH after registration.'
}
$wkhtmltopdfPath = $wkhtmltopdfCommand.Source
$ghostscriptCommand = Get-Command gswin64c.exe -ErrorAction SilentlyContinue
if (-not $ghostscriptCommand) {
    throw 'Ghostscript is not resolvable from PATH after registration.'
}
$ghostscriptPath = $ghostscriptCommand.Source

$wkhtmltopdfVersion = & $wkhtmltopdfPath --version 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "wkhtmltopdf was found at $wkhtmltopdfPath but failed its launch test."
}
$ghostscriptVersion = & $ghostscriptPath -version 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Ghostscript was found at $ghostscriptPath but failed its launch test."
}

$reg = Get-Command reg.exe -ErrorAction Stop
& $reg.Source import $registryFile
if ($LASTEXITCODE -ne 0) {
    throw "Failed to import printerRegister.reg (exit code $LASTEXITCODE)."
}

$protocolCommand = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey('spietopdf\shell\open\command', $true)
if (-not $protocolCommand) {
    throw 'The spietopdf URL protocol command key was not created by the registry import.'
}
try {
    $expectedProtocolCommand = "`"$batchFile`" `"%1`""
    $protocolCommand.SetValue('', $expectedProtocolCommand, [Microsoft.Win32.RegistryValueKind]::String)
    if ($protocolCommand.GetValue('') -ne $expectedProtocolCommand) {
        throw 'The spietopdf URL protocol command did not match the installed pdf.bat path.'
    }
}
finally {
    $protocolCommand.Dispose()
}

Write-Host 'PDF report dependencies are installed.'
Write-Host "wkhtmltopdf verified: $wkhtmltopdfPath ($($wkhtmltopdfVersion -join ' '))"
Write-Host "Ghostscript verified: $ghostscriptPath ($($ghostscriptVersion -join ' '))"
Write-Host "spietopdf: URL protocol registered to $batchFile"