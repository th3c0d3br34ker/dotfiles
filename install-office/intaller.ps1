<#
  .SYNOPSIS
    Installs Microsoft Office 365
  .DESCRIPTION
    Installs Microsoft Office 365 using a default configuration xml, unless a custom xml is provided.
    WARNING: This script will remove all existing office installations if used with the default configuration xml.
  .PARAMETER Config
    File path to custom configuration xml for office installations.
  .PARAMETER Cleanup
    Removes office installation files after install.
  .LINK
    XML Configuration Generator: https://config.office.com/
#>

param (
  [Alias('Configure')][String]$Config, # File path to custom configuration xml
  [Switch]$Cleanup # Cleans up installation files
)

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
  try { $ScriptPath = Convert-Path -Path $MyInvocation.MyCommand.Definition }
  catch {
    Write-Warning "Unable to resolve the script path for elevation: $_"
    exit 1
  }

  Write-Output 'Requesting administrator privileges...'

  $ElevatedArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
  if ($Config) { $ElevatedArgs += @('-Config', $Config) }
  if ($Cleanup) { $ElevatedArgs += '-Cleanup' }

  try {
    Start-Process -FilePath 'powershell.exe' -ArgumentList $ElevatedArgs -Verb RunAs | Out-Null
  }
  catch {
    Write-Warning 'Elevation request was cancelled or failed.'
    Write-Warning $_
    exit 1
  }

  exit 0
}

$ODT = "$env:temp\ODT"
$ConfigFile = "$ODT\office-config.xml"
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$LocalConfig = if ($ScriptDirectory) { Join-Path -Path $ScriptDirectory -ChildPath 'office-configuration.xml' } else { $null }
$LocalInstaller = if ($ScriptDirectory) { Join-Path -Path $ScriptDirectory -ChildPath 'ODTSetup.exe' } else { $null }
$InstallerFromTemp = "$env:temp\ODTSetup.exe"
$InstallerDownloaded = $false
$Installer = if ($LocalInstaller -and (Test-Path $LocalInstaller)) { $LocalInstaller } else { $InstallerFromTemp }
  
function Set-ConfigXML {
  param (
    [Parameter (Mandatory = $true)]
    [String]$XMLFile
  )
  
  $Path = Split-Path -Path $XMLFile -Parent
  if (!(Test-Path -PathType Container $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }

  $XML = [XML]@'
  <Configuration ID="5cf809c5-8f36-4fea-a837-69c7185cca8a">
    <Remove All="TRUE"/>
    <Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
      <Product ID="O365BusinessRetail">
        <Language ID="en-us"/>
        <ExcludeApp ID="Groove"/>
        <ExcludeApp ID="Lync"/>
        <ExcludeApp ID="Teams"/>
      </Product>
    </Add>
    <Property Name="SharedComputerLicensing" Value="0"/>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
    <Property Name="DeviceBasedLicensing" Value="0"/>
    <Property Name="SCLCacheOverride" Value="0"/>
    <Updates Enabled="TRUE"/>
    <RemoveMSI/>
    <AppSettings>
      <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas"/>
      <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas"/>
      <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas"/>
    </AppSettings>
    <Display Level="Full" AcceptEULA="TRUE"/>
  </Configuration>
'@

  $XML.Save("$XMLFile")
}

function Get-ODTURL {
  [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'
  $MSWebPage | ForEach-Object {
    if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') { $Matches[1] }
  }
}

# Set Config XML
if ($Config) {
  if (Test-Path $Config) { $ConfigFile = $Config }
  else {
    Write-Warning 'The configuration XML file path is not valid or is inaccessible.'
    Write-Warning 'Please check the path and try again.'
    exit 1
  }
}
elseif ($LocalConfig -and (Test-Path $LocalConfig)) {
  # Prefer a configuration file that ships with the installer when present
  $ConfigFile = $LocalConfig
}
else {
  Set-ConfigXML -XMLFile $ConfigFile
}

# Download Office Deployment Tool
if ($Installer -eq $InstallerFromTemp) {
  Write-Output 'Downloading Office Deployment Tool (ODT)...'
  $InstallLink = Get-ODTURL
  try {
    Invoke-WebRequest -Uri $InstallLink -OutFile $Installer
    $InstallerDownloaded = $true
  }
  catch {
    Write-Warning 'There was an error downloading the Office Deployment Tool.'
    Write-Warning 'Please verify the below link is valid:'
    Write-Warning $InstallLink
    exit 1
  }
}
else {
  Write-Output 'Using local Office Deployment Tool (ODT) installer...'
}

# Run Office Deployment Tool Setup
Write-Output 'Extracting Office Deployment Tool (ODT)...'
try { Start-Process -Wait -NoNewWindow -FilePath $Installer -ArgumentList "/extract:$ODT /quiet" }
catch {
  Write-Warning 'Error extracting Office Deployment Tool:'
  Write-Warning $_
  exit 1
}

# Install Office 
Write-Output 'Installing Microsoft Office...'
try { Start-Process -Wait -WindowStyle Hidden -FilePath "$ODT\setup.exe" -ArgumentList "/configure $ConfigFile" }
catch {
  Write-Warning 'Error during Office installation:'
  Write-Warning $_
  exit 1
}

# Check if Office 365 was installed
$RegPaths = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

foreach ($Key in (Get-ChildItem $RegPaths) ) {
  if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
    $InstalledVersion = $Key.GetValue('DisplayName')
  }
}

if ($InstalledVersion) { Write-Output "$InstalledVersion installed successfully!" }
else { Write-Warning 'Microsoft 365 was not detected after the installation completed. This warning is expected for non-365 office installs.' }

# Remove Office Hub
$AppName = 'Microsoft.MicrosoftOfficeHub'
try {
  Write-Output "Removing [$AppName] (Microsoft Store App)..."
  Get-AppxProvisionedPackage -Online | Where-Object { ($AppName -contains $_.DisplayName) } | Remove-AppxProvisionedPackage -AllUsers | Out-Null
  Get-AppxPackage -AllUsers | Where-Object { ($AppName -contains $_.Name) } | Remove-AppxPackage -AllUsers
}
catch { 
  Write-Warning "Error during [$AppName] removal:"
  Write-Warning $_
}

# Cleanup
if ($Cleanup) {
  Remove-Item $ODT -Recurse -Force -ErrorAction Ignore
  if ($InstallerDownloaded -and (Test-Path $Installer)) {
    Remove-Item $Installer -Force -ErrorAction Ignore
  }
}
