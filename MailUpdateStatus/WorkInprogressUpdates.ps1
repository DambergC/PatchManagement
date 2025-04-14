<#
.SYNOPSIS
    Extracts Windows Updates from a Configuration Manager Software Update Group and generates an HTML report.
.DESCRIPTION
    This script retrieves Windows updates from a specified SCCM Software Update Group,
    processes the update information, and creates an HTML report with conditional formatting
    for critical updates. It includes error handling, logging, and additional report features.
.NOTES
    File Name      : UpcomingUpdates.ps1
    Author         : DambergC
    Last Modified  : 2025-04-14
.PARAMETER UpdateGroupName
    The name of the software update group to process.
.PARAMETER HTMLReportPath
    The file path for saving the HTML report.
.PARAMETER LogFile
    The file path for saving the log file.
.PARAMETER SiteCode
    The site code of the Configuration Manager instance.
.PARAMETER ProviderMachineName
    The name of the provider machine in the Configuration Manager instance.
.PARAMETER ExportCSV
    Enables exporting updates to a CSV file.
.PARAMETER DryRun
    Simulates execution without making actual changes.
.EXAMPLE
    .\UpcomingUpdates.ps1 -UpdateGroupName "Server Patch Tuesday" -DryRun
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$UpdateGroupName,
    [string]$HTMLReportPath = "C:\temp\UpcomingUpdates.html",
    [string]$LogFile = "C:\temp\UpdateScript.log",
    [string]$SiteCode = "VP1",
    [string]$ProviderMachineName = "VXOVM20701",
    [switch]$ExportCSV = $true,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$CurrentDateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$CurrentUser = $env:USERNAME

# Enhanced Logging Function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogString,
        [string]$LogFile = $LogFile,
        [ValidateSet("Info", "Warning", "Error", "Verbose")]
        [string]$LogLevel = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$LogLevel] - $LogString"

    # Ensure directory exists
    $logDir = Split-Path -Path $LogFile -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    # Write to file
    Add-Content -Path $LogFile -Value $logEntry

    # Output to console
    switch ($LogLevel) {
        "Info" { Write-Host $logEntry -ForegroundColor Green }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Error" { Write-Host $logEntry -ForegroundColor Red }
        "Verbose" { Write-Verbose $logEntry }
    }
}

# Function to Connect to Configuration Manager
function Connect-ConfigurationManager {
    param (
        [string]$SiteCode,
        [string]$ProviderMachineName
    )
    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        Write-Log -LogString "Connecting to Configuration Manager site $SiteCode on $ProviderMachineName..." -LogLevel "Info"
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName -ErrorAction Stop
    }
    Push-Location
    Set-Location "$($SiteCode):\" -ErrorAction Stop
}

# Dry Run Mode
if ($DryRun) {
    Write-Log -LogString "Dry run mode enabled. No changes will be made." -LogLevel "Warning"
}

# Ensure the destination directory exists
try {
    $reportDirectory = Split-Path -Path $HTMLReportPath -Parent
    if (-not (Test-Path -Path $reportDirectory)) {
        if (-not $DryRun) {
            New-Item -Path $reportDirectory -ItemType Directory -Force -ErrorAction Stop
        }
        Write-Log -LogString "Created directory $reportDirectory" -LogLevel "Info"
    }
} catch {
    Write-Error "Failed to create directory $reportDirectory : $_"
    exit 1
}

# Log the beginning of the script
Write-Log -LogString "======================= Script START =============================" -LogLevel "Info"
Write-Log -LogString "Executed by: $CurrentUser at $CurrentDateTime" -LogLevel "Info"

try {
    # Connect to Configuration Manager
    Connect-ConfigurationManager -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName

    # Verify and Load PSWriteHTML Module
    if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
        Write-Log -LogString "PSWriteHTML module is not installed. Attempting to install..." -LogLevel "Warning"
        if (-not $DryRun) {
            Install-Module -Name PSWriteHTML -Force -ErrorAction Stop
        }
    }
    Import-Module PSWriteHTML -ErrorAction Stop
    Write-Log -LogString "PSWriteHTML module loaded successfully." -LogLevel "Info"

    # Retrieve Updates
    Write-Log -LogString "Retrieving updates from Software Update Group $UpdateGroupName..." -LogLevel "Info"
    if (-not $DryRun) {
        $updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName -ErrorAction Stop
    }

    # Process Updates
    $result = @()
    foreach ($item in $updates) {
        $object = [PSCustomObject]@{
            ArticleID            = $item.ArticleID
            Title                = $item.LocalizedDisplayName
            LocalizedDescription = $item.LocalizedDescription
            DatePosted           = $item.DatePosted.ToString("yyyy-MM-dd")
            Deployed             = $item.IsDeployed
            URL                  = $item.LocalizedInformativeURL
            Severity             = $item.SeverityName
        }
        $result += $object
    }

    # Generate HTML Report
    if (-not $DryRun) {
        New-HTML -FilePath $HTMLReportPath -Online -Title "Upcoming Windows Updates" {
            New-HTMLHeader {
                New-HTMLText -Text "Windows Updates Report - $UpdateGroupName" -Color Navy -Alignment center -FontSize 24
            }
            New-HTMLSection -HeaderText "Update Summary" -HeaderBackgroundColor Navy {
                New-HTMLText -Text "Total Updates: $($result.Count)" -Color Black -FontWeight Bold -FontSize 20
            }
            New-HTMLTable -DataTable $result -DisablePaging
        }
    }
    Write-Log -LogString "HTML report generated at $HTMLReportPath" -LogLevel "Info"

    # Export CSV (if enabled)
    if ($ExportCSV -and -not $DryRun) {
        $CSVReportPath = $HTMLReportPath.Replace(".html", ".csv")
        $result | Export-Csv -Path $CSVReportPath -NoTypeInformation -Force
        Write-Log -LogString "CSV report generated at $CSVReportPath" -LogLevel "Info"
    }

} catch {
    Write-Log -LogString "An error occurred: $_" -LogLevel "Error"
    exit 1
} finally {
    Pop-Location
    Write-Log -LogString "======================= Script END =============================" -LogLevel "Info"
}

# Output completion message
Write-Output "Script completed. Reports generated at:"
Write-Output "HTML: $HTMLReportPath"
if ($ExportCSV) { Write-Output "CSV: $CSVReportPath" }
Write-Output "Log: $LogFile"
