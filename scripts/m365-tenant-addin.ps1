<#
.SYNOPSIS
  Manage tenant-level Office Add-in deployment through Microsoft 365 Centralized Deployment.

.DESCRIPTION
  This script wraps Microsoft's tenant Office Add-in deployment cmdlets for the generated
  add-in-only XML manifests in dist/package/<profile>/centralized. The default lane contains
  one multi-host Word/Excel/PowerPoint manifest and one separate Outlook manifest.

  The script can use either ExchangeOnlineManagement or O365CentralizedAddInDeployment, but it
  selects by cmdlet availability instead of module name. Some current ExchangeOnlineManagement
  installs connect successfully yet do not expose New-OrganizationAddIn.

  It is intended for Microsoft 365 / Office on the web tenant testing where manually uploading a
  manifest in every browser is too slow. It does not replace the Integrated Apps portal for unified
  manifest packages.
#>

[CmdletBinding()]
param(
  [ValidateSet("List", "Deploy", "Update", "Delete", "Assign", "Enable", "Disable")]
  [string]$Action = "List",

  [string]$Profile = "development",
  [string]$ManifestPath,
  [string]$ManifestDir,
  [string]$ProductId,
  [string]$Locale = "en-US",
  [string[]]$Members = @(),
  [ValidateSet("Auto", "ExchangeOnlineManagement", "O365CentralizedAddInDeployment")]
  [string]$Backend = "Auto",
  [string]$UserPrincipalName,

  [switch]$AssignToEveryone,
  [switch]$UploadOnly,
  [switch]$InstallModule,
  [switch]$Device,
  [switch]$NoConnect,
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

function Write-Step {
  param([string]$Text)
  Write-Host ""
  Write-Host "==> $Text"
}

function Get-BackendModuleName {
  param([string]$Name)

  if ($Name -eq "ExchangeOnlineManagement") {
    return "ExchangeOnlineManagement"
  }
  return "O365CentralizedAddInDeployment"
}

function Test-BackendHasOrganizationAddInCmdlets {
  param([string]$Name)

  $moduleName = Get-BackendModuleName $Name
  if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    return $false
  }

  try {
    Import-Module -Name $moduleName -Force -ErrorAction Stop
  } catch {
    $script:BackendProbeErrors += "$moduleName import failed: $($_.Exception.Message)"
    return $false
  }

  return [bool](Get-Command -Name New-OrganizationAddIn -ErrorAction SilentlyContinue)
}

function Get-ResolvedBackend {
  if ($Backend -ne "Auto") {
    return $Backend
  }

  $script:BackendProbeErrors = @()
  foreach ($candidate in @("O365CentralizedAddInDeployment", "ExchangeOnlineManagement")) {
    if (Test-BackendHasOrganizationAddInCmdlets $candidate) {
      return $candidate
    }
  }

  if (Get-Module -ListAvailable -Name O365CentralizedAddInDeployment) {
    return "O365CentralizedAddInDeployment"
  }

  if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    return "ExchangeOnlineManagement"
  }

  return "O365CentralizedAddInDeployment"
}

function Ensure-CentralizedDeploymentModule {
  $script:ResolvedBackend = Get-ResolvedBackend
  $moduleName = Get-BackendModuleName $script:ResolvedBackend

  $module = Get-Module -ListAvailable -Name $moduleName
  if (-not $module) {
    if (-not $InstallModule) {
      throw "PowerShell module '$moduleName' is not installed. Re-run with -InstallModule or run: Install-Module -Name $moduleName -Scope CurrentUser"
    }
    Write-Step "Install $moduleName"
    Install-Module -Name $moduleName -Scope CurrentUser -Force
  }
  Import-Module -Name $moduleName -Force
}

function Assert-OrganizationAddInCmdlets {
  if (Get-Command -Name New-OrganizationAddIn -ErrorAction SilentlyContinue) {
    return
  }

  $moduleName = Get-BackendModuleName $script:ResolvedBackend
  $available = @(Get-Command -Name "*OrganizationAddIn*" -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }) -join ", "
  if (-not $available) {
    $available = "none"
  }

  throw @"
Backend '$script:ResolvedBackend' loaded module '$moduleName', but New-OrganizationAddIn is unavailable.
Available OrganizationAddIn commands: $available

Use -Backend Auto to let the script select a working backend, or install the Centralized Deployment module:
  Install-Module -Name O365CentralizedAddInDeployment -Scope CurrentUser -Force
"@
}

function Connect-IfNeeded {
  if ($NoConnect) {
    return
  }

  if ($script:ResolvedBackend -eq "ExchangeOnlineManagement") {
    Write-Step "Connect to Exchange Online Management"
    $params = @{
      ShowBanner = $false
    }
    if ($UserPrincipalName) {
      $params.UserPrincipalName = $UserPrincipalName
    }
    if ($Device) {
      $params.Device = $true
    }
    Connect-ExchangeOnline @params
  } else {
    Write-Step "Connect to Microsoft 365 Centralized Deployment"
    try {
      Connect-OrganizationAddInService
    } catch {
      $message = $_.Exception.Message
      if ($IsLinux -and $message -match "kernel32\.dll") {
        throw @"
O365CentralizedAddInDeployment exposes New-OrganizationAddIn here, but its Connect-OrganizationAddInService path is loading Windows-native kernel32.dll and cannot authenticate under this Linux PowerShell runtime.

This machine can still package manifests and run -NoConnect -DryRun, but live Centralized Deployment requires one of:
  1. Run this same script from PowerShell on Windows with -Backend Auto.
  2. Use the Microsoft 365 admin center Integrated Apps / Centralized Deployment portal.
  3. Re-test later with -Backend ExchangeOnlineManagement if Microsoft's ExchangeOnlineManagement module starts exposing New-OrganizationAddIn in your tenant.

Original error: $message
"@
      }
      throw
    }
  }
}

function Resolve-ManifestPaths {
  if ($ManifestPath) {
    return @(Resolve-Path $ManifestPath | ForEach-Object { $_.Path })
  }

  $dir = $ManifestDir
  if (-not $dir) {
    $dir = Join-Path $RepoRoot "dist/package/$Profile/centralized"
  }
  if (-not (Test-Path $dir)) {
    throw "Manifest directory does not exist: $dir. Run bun run package:dev first."
  }

  $files = Get-ChildItem -Path $dir -Filter "*.manifest.xml" -File | Sort-Object Name
  if (-not $files) {
    throw "No *.manifest.xml files found in $dir"
  }
  return @($files | ForEach-Object { $_.FullName })
}

function Get-ManifestProductId {
  param([string]$Path)

  $node = Select-Xml -Path $Path -XPath "/*[local-name()='OfficeApp']/*[local-name()='Id'][1]" -ErrorAction SilentlyContinue
  if (-not $node) {
    $node = Select-Xml -Path $Path -XPath "//*[local-name()='Id'][1]" -ErrorAction SilentlyContinue
  }
  if (-not $node) {
    throw "Could not find manifest Id in $Path"
  }
  return [string]$node.Node.InnerText
}

function Get-EffectiveProductIds {
  if ($ProductId) {
    return @($ProductId)
  }
  return @(Resolve-ManifestPaths | ForEach-Object { Get-ManifestProductId $_ })
}

function Invoke-OrPrint {
  param(
    [string]$Description,
    [scriptblock]$Command
  )
  if ($DryRun) {
    Write-Host "[dry-run] $Description"
    return $null
  }
  Write-Host $Description
  return & $Command
}

function Test-OrgAddInExists {
  param([string]$Id)
  try {
    $all = Get-OrganizationAddIn -ErrorAction SilentlyContinue
  } catch {
    return $false
  }
  return @($all | Where-Object { $_.ProductId -eq $Id }).Count -gt 0
}

function Add-AssignmentsIfRequested {
  param([string]$EffectiveProductId)

  if ($AssignToEveryone) {
    Invoke-OrPrint `
      "Set-OrganizationAddInAssignments -ProductId $EffectiveProductId -AssignToEveryone `$true" `
      { Set-OrganizationAddInAssignments -ProductId $EffectiveProductId -AssignToEveryone $true }
  }

  if ($Members.Count -gt 0) {
    $memberText = $Members -join ", "
    Invoke-OrPrint `
      "Set-OrganizationAddInAssignments -ProductId $EffectiveProductId -Add -Members $memberText" `
      { Set-OrganizationAddInAssignments -ProductId $EffectiveProductId -Add -Members $Members }
  }
}

Ensure-CentralizedDeploymentModule
Connect-IfNeeded
Assert-OrganizationAddInCmdlets

switch ($Action) {
  "List" {
    Write-Step "List deployed organization add-ins"
    Invoke-OrPrint "Get-OrganizationAddIn" { Get-OrganizationAddIn | Format-Table -AutoSize }
  }

  "Deploy" {
    if (-not $UploadOnly -and -not $AssignToEveryone -and $Members.Count -eq 0) {
      throw "Deploy requires -UploadOnly, -AssignToEveryone, or -Members. This prevents accidental tenant-wide ambiguity."
    }

    $manifests = Resolve-ManifestPaths
    Write-Step "Deploy $($manifests.Count) manifest(s)"
    foreach ($manifest in $manifests) {
      $manifestProductId = Get-ManifestProductId $manifest
      # Idempotent upsert: if this ProductId is already deployed, update it in place instead of
      # letting New-OrganizationAddIn fail on a duplicate. This makes `deploy` safe to re-run and
      # robust to a lost .ge-deploy.json state file on an already-provisioned tenant.
      if (Test-OrgAddInExists $manifestProductId) {
        Invoke-OrPrint `
          "Set-OrganizationAddIn -ProductId $manifestProductId -ManifestPath '$manifest' -Locale '$Locale' (already deployed)" `
          { Set-OrganizationAddIn -ProductId $manifestProductId -ManifestPath $manifest -Locale $Locale }
      } elseif ($Members.Count -gt 0) {
        $memberText = $Members -join ", "
        Invoke-OrPrint `
          "New-OrganizationAddIn -ManifestPath '$manifest' -Locale '$Locale' -Members $memberText" `
          { New-OrganizationAddIn -ManifestPath $manifest -Locale $Locale -Members $Members }
      } else {
        Invoke-OrPrint `
          "New-OrganizationAddIn -ManifestPath '$manifest' -Locale '$Locale'" `
          { New-OrganizationAddIn -ManifestPath $manifest -Locale $Locale }
      }

      if (-not $UploadOnly) {
        Add-AssignmentsIfRequested -EffectiveProductId $manifestProductId
      }
    }
  }

  "Update" {
    $manifests = Resolve-ManifestPaths
    Write-Step "Update $($manifests.Count) manifest(s)"
    foreach ($manifest in $manifests) {
      $effectiveProductId = if ($ProductId) { $ProductId } else { Get-ManifestProductId $manifest }
      Invoke-OrPrint `
        "Set-OrganizationAddIn -ProductId $effectiveProductId -ManifestPath '$manifest' -Locale '$Locale'" `
        { Set-OrganizationAddIn -ProductId $effectiveProductId -ManifestPath $manifest -Locale $Locale }
      Add-AssignmentsIfRequested -EffectiveProductId $effectiveProductId
    }
  }

  "Delete" {
    $ids = Get-EffectiveProductIds
    Write-Step "Delete $($ids.Count) add-in(s)"
    foreach ($id in $ids) {
      Invoke-OrPrint "Remove-OrganizationAddIn -ProductId $id" { Remove-OrganizationAddIn -ProductId $id }
    }
  }

  "Assign" {
    if (-not $AssignToEveryone -and $Members.Count -eq 0) {
      throw "Assign requires -AssignToEveryone or -Members."
    }
    $ids = Get-EffectiveProductIds
    Write-Step "Assign $($ids.Count) add-in(s)"
    foreach ($id in $ids) {
      Add-AssignmentsIfRequested -EffectiveProductId $id
    }
  }

  "Enable" {
    $ids = Get-EffectiveProductIds
    Write-Step "Enable $($ids.Count) add-in(s)"
    foreach ($id in $ids) {
      Invoke-OrPrint "Set-OrganizationAddIn -ProductId $id -Enabled `$true" { Set-OrganizationAddIn -ProductId $id -Enabled $true }
    }
  }

  "Disable" {
    $ids = Get-EffectiveProductIds
    Write-Step "Disable $($ids.Count) add-in(s)"
    foreach ($id in $ids) {
      Invoke-OrPrint "Set-OrganizationAddIn -ProductId $id -Enabled `$false" { Set-OrganizationAddIn -ProductId $id -Enabled $false }
    }
  }
}
