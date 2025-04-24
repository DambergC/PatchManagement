<#
.SYNOPSIS
    Update status for Windows update through MECM.

.DESCRIPTION
    This script generates a report on Windows updates for specified deployments
    and sends an email to the configured recipients. It supports manual and scheduled execution.

.NOTES
    Created by: Christian Damberg
    Organization: Telia Cygate AB
    Filename: Send-WindowsUpdateStatus.ps1

    Prerequisites:
    - Required PowerShell Modules: send-mailkitmessage, PSWriteHTML
    - Configuration file: ScriptConfig.xml
    - Ensure proper permissions to access all required resources.

    Usage:
    - Run the script manually or schedule it with Task Scheduler.
    - Ensure the configuration file and required modules are in place before execution.
#>

param (
    [string]$ConfigFilePath = ".\ScriptConfig.xml",
    [switch]$UseSecureConnection = $false
)

try {
    if (-not (Test-Path $ConfigFilePath)) {
        throw "Configuration file not found at $ConfigFilePath."
    }

    $xml = [System.Xml.XmlDocument]::new()
    $xml.Load($ConfigFilePath)

    # Validate Configuration
    if (-not $xml.Configuration) {
        throw "Missing 'Configuration' section in XML."
    }
    if (-not $xml.Configuration.Logfile.Path -or -not $xml.Configuration.MailSMTP) {
        throw "Missing required XML configuration nodes."
    }

} catch {
    Write-Host "Error: $_.Exception.Message"
    exit 1
}

# Logging Setup
$Logfilepath = $xml.Configuration.Logfile.Path
$logfilename = $xml.Configuration.Logfile.Name
$Logfile = Join-Path $Logfilepath $logfilename
$Logfilethreshold = $xml.Configuration.Logfile.Logfilethreshold

function Write-Log {
    param (
        [String]$LogString
    )
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    try {
        Add-Content $LogFile -Value $LogMessage
    } catch {
        Write-Host "Failed to write log: $_"
    }
}

function Rotate-Log {
    $target = Get-ChildItem -Path $Logfilepath -Filter "windo*.log"
    $datetime = Get-Date -uformat "%Y-%m-%d-%H%M"

    foreach ($file in $target) {
        if ($file.Length -ge $Logfilethreshold) {
            try {
                $newName = "$($file.BaseName)_${datetime}.log"
                $oldLogPath = Join-Path $Logfilepath "OLDLOG"

                if (-not (Test-Path $oldLogPath)) {
                    New-Item -Path $oldLogPath -ItemType Directory | Out-Null
                }

                Move-Item -Path $file.FullName -Destination $oldLogPath
                Compress-Archive -Path "$oldLogPath\$($file.Name)" -DestinationPath "$oldLogPath\$($file.BaseName).zip"
                Remove-Item "$oldLogPath\$($file.Name)"
                Write-Log -LogString "Compressed and archived log file: $($file.Name)"
            } catch {
                Write-Log -LogString "Failed to rotate log file $($file.Name): $_"
            }
        } else {
            Write-Log -LogString "Log file $($file.Name) does not need rotation."
        }
    }
}

Rotate-Log

# Dependency Management
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$RequiredVersion = "0.0.0"
    )
    $Module = Get-Module -Name $ModuleName -ListAvailable
    if (-not $Module -or $Module.Version -lt $RequiredVersion) {
        try {
            Install-Module -Name $ModuleName -RequiredVersion $RequiredVersion -Force
        } catch {
            Write-Log -LogString "Error installing module $ModuleName: $_.Exception.Message"
        }
    }
    Import-Module -Name $ModuleName
}

Ensure-Module -ModuleName "send-mailkitmessage"
Ensure-Module -ModuleName "PSWriteHTML"

# Secure Credential Handling
$Credential = Get-StoredCredential -Target "SMTP_Credentials"
if (-not $Credential) {
    Write-Log -LogString "Error: SMTP credentials not found in Credential Manager."
    exit 1
}

# Site Server and Patch Tuesday Handling
function Get-CMSiteCode {
    $CMSiteCode = Get-CimInstance -Namespace "root\SMS" -ClassName SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
    return $CMSiteCode
}

$todayDefault = Get-Date
$todayshort = $todayDefault.ToShortDateString()
$thismonth = $todayDefault.Month
$nextmonth = $todayDefault.AddMonths(1).Month
$patchtuesdayThisMonth = Get-PatchTuesday -Month $thismonth -Year $todayDefault.Year
$patchtuesdayNextMonth = Get-PatchTuesday -Month $nextmonth -Year $todayDefault.Year

# Deployment Processing
$ParametersNode = $xml.Configuration.Runscript.SelectNodes('//Job')
foreach ($Node in $ParametersNode) {
    # Processing Details
    $HTMLfilepath = $xml.Configuration.HTMLfilePath
    $DescToFile = $Node.Description
    $filedate = Get-Date -Format "yyyyMMdd"
    $filename = "${DescToFile}_UpdateStatus_${filedate}.HTML"
    $HTMLFileSavePath = Join-Path $HTMLfilepath $filename

    $deploymentIDtoCheck = $Node.DeploymentID
    $DaysAfterPatchTuesdayToReport = $Node.Offsetdays
    $ReportdayCompare = ($patchtuesdayThisMonth.AddDays($DaysAfterPatchTuesdayToReport)).ToString("yyyy-MM-dd")

    # Collect Deployment Data
    $ResultColl = @()

    if ($todayshort -eq $ReportdayCompare) {
        Write-Log -LogString "Collecting data for deployment ID $deploymentIDtoCheck"
        $UpdateStatus = Get-SCCMSoftwareUpdateStatus -DeploymentID $deploymentIDtoCheck

        foreach ($US in $UpdateStatus) {
            $object = [PSCustomObject]@{
                Server       = $US.DeviceName
                Collection   = $US.CollectionName
                Status       = $US.Status
                StatusTime   = $US.StatusTime
            }
            $ResultColl += $object
        }
    } else {
        Write-Log -LogString "Date mismatch: Expected $ReportdayCompare, but today is $todayshort. Skipping."
        continue
    }

    # Generate HTML Report
    New-HTML -TitleText "Update Status - $MailCustomer" -FilePath $HTMLFileSavePath -ShowHTML -Online {
        New-HTMLSection -Invisible -Title "Summary" {
            New-HTMLTable -DataTable @(
                @{ Key = "Total Servers"; Value = $ResultColl.Count },
                @{ Key = "Success"; Value = ($UpdateStatus | Where-Object { $_.Status -eq 'success' }).Count },
                @{ Key = "Needs Attention"; Value = ($UpdateStatus | Where-Object { $_.Status -ne 'success' }).Count }
            ) -Style compact
        }
    }

    # Send Email
    $Parameters = @{
        "SMTPServer"                     = $SMTP
        "Port"                           = $MailPortnumber
        "From"                           = $MailFrom
        "RecipientList"                  = $xml.Configuration.Recipients | ForEach-Object { $_.Recipients.Email }
        "Subject"                        = "Update Status Report - $DescToFile"
        "HTMLBody"                       = Get-Content $HTMLFileSavePath -Raw
        "AttachmentList"                 = @($HTMLFileSavePath)
    }

    Send-MailKitMessage @Parameters
    Write-Log -LogString "Email sent for deployment ID $deploymentIDtoCheck"
}

Write-Log -LogString "Script execution completed."
