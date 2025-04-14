<#
.SYNOPSIS
    Extracts Windows Updates from a Configuration Manager Software Update Group and generates an HTML report.
.DESCRIPTION
    This script retrieves Windows updates from a specified SCCM Software Update Group,
    processes the update information, and creates an HTML report with conditional formatting
    for critical updates. It includes error handling, logging, and additional report features.
.NOTES
    File Name      : Windows_Update_Report.ps1
    Author         : DambergC
    Last Modified  : 2025-04-14
#>

# Script Name
$scriptname = "Windows Update Extraction with HTML Report"

# Variables
$UpdateGroupName = "Server Patch Tuesday" # Replace with the name of your Software Update Group
$HTMLReportPath = "C:\temp\UpcomingUpdates.html" # Path to save the HTML report
$LogFile = "C:\temp\UpdateScript.log" # Path to save the log file
$SiteCode = "VP1" # Replace with your site code
$ProviderMachineName = "VXOVM20701" # Replace with your server name
$ExportCSV = $true # Set to $false if you don't want a CSV export
$CSVReportPath = $HTMLReportPath.Replace(".html", ".csv")

# Define Write-Log function if not already available
if (-not (Get-Command 'Write-Log' -ErrorAction SilentlyContinue)) {
    # Replace the entire Write-Log function with this more robust version
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogString,
        [string]$LogFile = $script:LogFile
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp - $LogString"
        
        # Ensure directory exists
        $logDir = Split-Path -Path $LogFile -Parent
        if (-not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Create or open the file with proper encoding
        if (-not (Test-Path -Path $LogFile)) {
            # Create new file with UTF-8 encoding (no BOM)
            $streamWriter = New-Object System.IO.StreamWriter($LogFile, $false, (New-Object System.Text.UTF8Encoding($false)))
        } else {
            # Append to existing file with UTF-8 encoding (no BOM)
            $streamWriter = New-Object System.IO.StreamWriter($LogFile, $true, (New-Object System.Text.UTF8Encoding($false)))
        }
        
        # Write the log entry and close the file
        $streamWriter.WriteLine($logEntry)
        $streamWriter.Close()
        $streamWriter.Dispose()
        
        # Output to console as well if verbose
        Write-Verbose $logEntry
    }
    catch {
        Write-Error "Failed to write to log file: $_"
    }
}
}

# Current date and user info for reporting
$CurrentDateTime = "2025-04-14 13:13:26" # UTC time as provided
$CurrentUser = "DambergC" # User's login as provided

# Ensure the destination directory exists
try {
    $reportDirectory = Split-Path -Path $HTMLReportPath -Parent
    if (-not (Test-Path -Path $reportDirectory)) {
        New-Item -Path $reportDirectory -ItemType Directory -Force -ErrorAction Stop
        Write-Log -LogString "Created directory $reportDirectory"
    }
} catch {
    Write-Error "Failed to create directory $reportDirectory : $_"
    exit 1
}

# Log the beginning of the script
Write-Log -LogString "======================= $scriptname - Script START ============================="
Write-Log -LogString "Executed by: $CurrentUser at $CurrentDateTime"

# Check if Configuration Manager module is loaded
try {
    if (-not (Get-Module ConfigurationManager)) {
        Write-Log -LogString "Importing Configuration Manager module..."
        Import-Module ConfigurationManager -ErrorAction Stop
    }
} catch {
    $errorMsg = "Failed to load Configuration Manager module: $_"
    Write-Log -LogString $errorMsg
    Write-Error $errorMsg
    exit 1
}

# Connect to Configuration Manager site if not already connected
try {
    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        Write-Log -LogString "Connecting to Configuration Manager site $SiteCode on $ProviderMachineName..."
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName -ErrorAction Stop
    }
    
    # Set location to the Configuration Manager site
    Push-Location
    Set-Location "$($SiteCode):\" -ErrorAction Stop
} catch {
    $errorMsg = "Failed to connect to Configuration Manager site: $_"
    Write-Log -LogString $errorMsg
    Write-Error $errorMsg
    exit 1
}

# Check if PSWriteHTML module is installed
try {
    if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
        Write-Log -LogString "Required module PSWriteHTML is not installed. Attempting to install..."
        Install-Module -Name PSWriteHTML -Force -ErrorAction Stop
        Write-Log -LogString "PSWriteHTML module installed successfully"
    }
    
    # Import the module
    Import-Module PSWriteHTML -ErrorAction Stop
    
    # Get module version for troubleshooting
    $moduleVersion = (Get-Module PSWriteHTML).Version
    Write-Log -LogString "Using PSWriteHTML module version $moduleVersion"
} catch {
    $errorMsg = "Error with PSWriteHTML module: $_"
    Write-Log -LogString $errorMsg
    Write-Error $errorMsg
    exit 1
}

# Get updates in the specified Software Update Group
try {
    Write-Log -LogString "===================== Processing Software Update Group $UpdateGroupName ==========================="
    $updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName -ErrorAction Stop
    
    if ($updates.Count -eq 0) {
        $warningMsg = "No updates found in update group '$UpdateGroupName'"
        Write-Log -LogString $warningMsg
        Write-Warning $warningMsg
    } else {
        Write-Log -LogString "Found $($updates.Count) updates in group '$UpdateGroupName'"
    }
} catch {
    $errorMsg = "Error retrieving updates: $_"
    Write-Log -LogString $errorMsg
    Write-Error $errorMsg
    
    # Return to original location before exiting
    Pop-Location
    exit 1
}

# Initialize result array
$result = @()

# Process each update
try {
    Write-Log -LogString "Processing updates..."
    foreach ($item in $updates) {
        $object = [PSCustomObject]@{
            ArticleID            = $item.ArticleID
            Title                = $item.LocalizedDisplayName
            LocalizedDescription = $item.LocalizedDescription
            DatePosted           = $item.DatePosted.ToString("yyyy-MM-dd") # Format to show date only
            Deployed             = $item.IsDeployed
            URL                  = $item.LocalizedInformativeURL
            Severity             = $item.SeverityName
        }
        $result += $object
    }
} catch {
    $errorMsg = "Error processing updates: $_"
    Write-Log -LogString $errorMsg
    Write-Error $errorMsg
}

# Calculate statistics
$totalUpdates = $result.Count
$criticalUpdates = ($result | Where-Object {$_.Severity -eq 'Critical'} | Measure-Object).Count
$highUpdates = ($result | Where-Object {$_.Severity -eq 'Important'} | Measure-Object).Count
$mediumUpdates = ($result | Where-Object {$_.Severity -eq 'Moderate'} | Measure-Object).Count
$lowUpdates = ($result | Where-Object {$_.Severity -eq 'Low'} | Measure-Object).Count

Write-Log -LogString "Update Summary - Total: $totalUpdates, Critical: $criticalUpdates, Important: $highUpdates, Moderate: $mediumUpdates, Low: $lowUpdates"

# Remove existing files if they exist to ensure overwrite
if (Test-Path -Path $HTMLReportPath) {
    Remove-Item -Path $HTMLReportPath -Force
    Write-Log -LogString "Removed existing HTML report file"
}

if ($ExportCSV -and (Test-Path -Path $CSVReportPath)) {
    Remove-Item -Path $CSVReportPath -Force
    Write-Log -LogString "Removed existing CSV report file"
}

# Generate HTML Report with New-HTMLSection
try {
    Write-Log -LogString "Generating HTML report at $HTMLReportPath..."
    
    # Create HTML report with PSWriteHTML syntax
    New-HTML -FilePath $HTMLReportPath -Online -Title "Upcoming Windows Updates" {
        New-HTMLHeader {
            New-HTMLText -Text "Windows Updates Report - $UpdateGroupName" -Color Navy -Alignment center -FontSize 24
            
        }

        # Add summary section
 

        New-HTMLSection -HeaderText "Total Updates" -HeaderBackgroundColor Navy -HeaderTextSize 30 {
            New-HTMLPanel {
                New-HTMLText -Text "$totalUpdates" -Color Black -FontWeight bold -FontSize 40 -Alignment center

            }

        }

            New-HTMLSection -HeaderText "Update Summary"  -HeaderBackgroundColor Navy -HeaderTextSize 30 {
            New-HTMLPanel {
                New-HTMLText -Text "Critical Updates: $criticalUpdates" -Color Red -FontWeight bold -FontSize 20 -Alignment center

            }
                        New-HTMLPanel {
                New-HTMLText -Text "Important Updates: $highUpdates" -Color Orange -FontWeight bold -FontSize 20 -Alignment center

            }
                        New-HTMLPanel {
                New-HTMLText -Text "Moderate Updates: $mediumUpdates" -Color Blue -FontWeight bold -FontSize 20 -Alignment center
            }
                        New-HTMLPanel {
                New-HTMLText -Text "Low Updates: $lowUpdates" -Color Green -FontWeight bold -FontSize 20 -Alignment center
            }
            }
        
        # Add table within its own section
        New-HTMLSection -HeaderText "Windows Updates Details" -HeaderBackgroundColor Navy -HeaderTextColor White -HeaderTextSize 30 {
            New-HTMLTable -DataTable $result -HideFooter -DisablePaging {
                New-TableCondition -Name 'Severity' -ComparisonType string -Operator eq -Value 'Critical' -BackgroundColor Red -Color White
                New-TableCondition -Name 'Severity' -ComparisonType string -Operator eq -Value 'Important' -BackgroundColor Orange
                New-TableCondition -Name 'Severity' -ComparisonType string -Operator eq -Value 'Moderate' -BackgroundColor LightBlue
                New-TableCondition -Name 'Severity' -ComparisonType string -Operator eq -Value 'Low' -BackgroundColor LightGreen
            }
        }
    }
    
    Write-Log -LogString "HTML Report generated successfully at $HTMLReportPath"
} catch {
    $errorMsg = "Error generating HTML report: $_"
    Write-Log -LogString $errorMsg
    Write-Error $errorMsg
}

# Export to CSV if enabled
if ($ExportCSV) {
    try {
        $result | Export-Csv -Path $CSVReportPath -NoTypeInformation -Force
        Write-Log -LogString "CSV Report generated at $CSVReportPath"
    } catch {
        $errorMsg = "Error generating CSV report: $_"
        Write-Log -LogString $errorMsg
        Write-Error $errorMsg
    }
}

# Return to original location
Pop-Location

# Log the end of the script
Write-Log -LogString "======================= $scriptname - Script END ============================="

# Show completion message
Write-Output "Script completed. Reports generated at:"
Write-Output "HTML: $HTMLReportPath"
if ($ExportCSV) { Write-Output "CSV: $CSVReportPath" }
Write-Output "Log: $LogFile"
