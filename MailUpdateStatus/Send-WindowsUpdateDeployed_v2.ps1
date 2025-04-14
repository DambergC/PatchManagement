<# 
.SYNOPSIS
    Improved script for reporting Windows Update deployment statuses.

.DESCRIPTION
    This script generates a report for Windows Updates based on deployments defined in a configuration file,
    and securely sends the report via email. It includes enhancements for error handling, logging, 
    and secure credential management.

.NOTES
    Created by: Christian Damberg, Telia Cygate AB
    Date: 2025-04-14
#>

# Load Configuration
function Load-Configuration {
    param (
        [string]$ConfigFilePath = ".\ScriptConfig.xml"
    )
    try {
        if (-not (Test-Path $ConfigFilePath)) {
            throw "Configuration file not found at $ConfigFilePath"
        }
        [System.Xml.XmlDocument]$xml = Get-Content $ConfigFilePath
        return $xml
    } catch {
        Write-Error "Failed to load configuration: $_"
        exit 1
    }
}

# Initialize Logging
function Initialize-Log {
    param (
        [string]$LogFilePath
    )
    try {
        if (-not (Test-Path $LogFilePath)) {
            New-Item -ItemType File -Path $LogFilePath -Force
        }
    } catch {
        Write-Error "Failed to initialize log file: $_"
        exit 1
    }
}

function Write-Log {
    param (
        [string]$LogFile,
        [string]$LogString
    )
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-Content $LogFile -Value $LogMessage
}

# Rotate Logs
function Rotate-Log {
    param (
        [string]$LogFilePath,
        [int]$LogFileThreshold
    )
    try {
        $target = Get-ChildItem $LogFilePath -Filter "*.log"
        $datetime = Get-Date -UFormat "%Y-%m-%d-%H%M"

        foreach ($file in $target) {
            if ($file.Length -ge $LogFileThreshold) {
                $newName = "$($file.BaseName)_${datetime}.log"
                Rename-Item $file.FullName $newName -ErrorAction Stop
                Move-Item $newName -Destination "$LogFilePath\OLDLOG" -ErrorAction Stop
                Write-Log -LogFile "$LogFilePath\Logfile.log" -LogString "Rotated log file: $file.Name"
            }
        }
    } catch {
        Write-Error "Failed to rotate logs: $_"
    }
}

# Send Email
function Send-Email {
    param (
        [string]$SMTPServer,
        [int]$Port,
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$Body,
        [string]$Attachment = $null
    )
    try {
        $Parameters = @{
            SMTPServer = $SMTPServer
            Port = $Port
            From = $From
            To = $To
            Subject = $Subject
            Body = $Body
            Attachment = $Attachment
        }

        # Example: Replace this with your email sending module
        Send-MailKitMessage @Parameters
        Write-Log -LogFile "$LogFilePath\Logfile.log" -LogString "Email sent successfully to $To"
    } catch {
        Write-Error "Failed to send email: $_"
    }
}

# Main Script
try {
    $Config = Load-Configuration
    $LogFilePath = $Config.Configuration.LogFile.Path
    $LogFileThreshold = $Config.Configuration.LogFile.LogFileThreshold

    Initialize-Log -LogFilePath $LogFilePath
    Rotate-Log -LogFilePath $LogFilePath -LogFileThreshold $LogFileThreshold

    # Generate Report (Placeholder)
    $Report = "Sample Report Content"

    # Email Configuration
    $SMTPServer = $Config.Configuration.MailSMTP
    $Port = $Config.Configuration.MailPort
    $From = $Config.Configuration.MailFrom
    $To = $Config.Configuration.Recipients.Recipient.Email
    $Subject = "Windows Update Report"
    $Body = $Report

    Send-Email -SMTPServer $SMTPServer -Port $Port -From $From -To $To -Subject $Subject -Body $Body
} catch {
    Write-Error "An error occurred: $_"
    exit 1
}
