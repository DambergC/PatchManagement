<#
    .SYNOPSIS
        Update status for Windows updates through MECM.

    .DESCRIPTION
        The script creates a report on Windows updates for one or more deployments
        specified in an XML file and sends an email to named recipients.
        The script can run manually or be scheduled on the site server.

    .NOTES
        ===========================================================================
        Created with:    SAPIEN Technologies, Inc., PowerShell Studio 2024
        Created on:      10/16/2023 3:34 PM
        Updated on:      04/14/2025
        Created by:      Christian Damberg
        Organization:    Telia Cygate AB
        Filename:        Send-WindowsUpdateDeployed.ps1
        Improvements:    Modularized, error-handling, logging levels, retry mechanisms, and best practices applied.
        ===========================================================================
#>

[System.Xml.XmlDocument]$xml = Get-Content .\ScriptConfig.xml

# Configuration Variables
$Logfilepath = $xml.Configuration.Logfile.Path
$logfilename = $xml.Configuration.Logfile.Name
$LogFile = Join-Path $Logfilepath $logfilename
$Logfilethreshold = $xml.Configuration.Logfile.Logfilethreshold
$SMTP = $xml.Configuration.MailSMTP
$MailFrom = $xml.Configuration.Mailfrom
$MailPortnumber = $xml.Configuration.MailPort
$MailCustomer = $xml.Configuration.MailCustomer
$LimitDays = $xml.Configuration.UpdateDeployed.LimitDays
$DaysAfterPatchTuesdayToReport = $xml.Configuration.UpdateDeployed.DaysAfterPatchToRun
$UpdateGroupName = $xml.Configuration.UpdateDeployed.UpdateGroupName

# Logging Function with Levels
function Write-Log {
    param (
        [string]$LogString,
        [ValidateSet("INFO", "WARN", "ERROR")] [string]$LogLevel = "INFO"
    )
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp [$LogLevel] $LogString"
    Add-Content $LogFile -Value $LogMessage
}

# Rotate Logs with Error Handling
function Rotate-Log {
    try {
        $target = Get-ChildItem $Logfilepath -Filter "windo*.log"
        $datetime = Get-Date -uformat "%Y-%m-%d-%H%M"

        foreach ($file in $target) {
            if ($file.Length -ge $Logfilethreshold) {
                Write-Log -LogString "Rotating log file: $($file.Name)" -LogLevel INFO
                $newname = "$($file.BaseName)_${datetime}.log"
                Rename-Item $file.FullName $newname
                $oldLogDir = Join-Path $Logfilepath "OLDLOG"
                if (-not (Test-Path $oldLogDir)) {
                    New-Item -Path $oldLogDir -ItemType Directory
                }
                Move-Item $newname -Destination $oldLogDir
            }
        }
    } catch {
        Write-Log -LogString "Error during log rotation: $_" -LogLevel ERROR
    }
}

# Retry Mechanism for Email Sending
function Send-EmailWithRetry {
    param (
        [hashtable]$Parameters,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 5
    )
    $attempt = 0
    do {
        try {
            Send-MailKitMessage @Parameters
            Write-Log -LogString "Email sent successfully on attempt $($attempt + 1)" -LogLevel INFO
            break
        } catch {
            $attempt++
            Write-Log -LogString "Failed to send email on attempt $attempt. Retrying... $_" -LogLevel WARN
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    } while ($attempt -lt $MaxRetries)

    if ($attempt -ge $MaxRetries) {
        Write-Log -LogString "Failed to send email after $MaxRetries attempts." -LogLevel ERROR
        throw "Email sending failed."
    }
}

# Import Required Modules with Error Handling
function Import-RequiredModules {
    $modules = @(
        "ConfigurationManager",
        "send-mailkitmessage",
        "PSWriteHTML",
        "PatchManagementSupportTools"
    )
    foreach ($module in $modules) {
        if (-not (Get-Module -Name $module)) {
            try {
                Import-Module $module -ErrorAction Stop
                Write-Log -LogString "Module $module imported successfully." -LogLevel INFO
            } catch {
                Write-Log -LogString "Failed to import module: $module. $_" -LogLevel ERROR
                throw
            }
        }
    }
}

# Main Script Execution
try {
    Write-Log -LogString "Starting script execution." -LogLevel INFO

    # Import Required Modules
    Import-RequiredModules

    # Rotate Logs
    Rotate-Log

    # Determine Report Day
    $patchTuesdayThisMonth = Get-PatchTuesday -Month (Get-Date).Month -Year (Get-Date).Year
    $ReportdayCompare = $patchTuesdayThisMonth.AddDays($DaysAfterPatchTuesdayToReport).ToString("yyyy-MM-dd")
    if ((Get-Date).ToString("yyyy-MM-dd") -eq $ReportdayCompare) {
        Write-Log -LogString "Script is running on the scheduled report day." -LogLevel INFO

        # Collect Data
        $updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName
        $result = @()
        foreach ($item in $updates) {
            $result += [PSCustomObject]@{
                ArticleID             = $item.ArticleID
                Title                 = $item.LocalizedDisplayName
                LocalizedDescription  = $item.LocalizedDescription
                DatePosted            = $item.DatePosted
                Deployed              = $item.IsDeployed
                URL                   = $item.LocalizedInformativeURL
                Severity              = $item.SeverityName
            }
        }

        # Prepare Email Parameters
        $EmailParams = @{
            "SMTPServer"    = $SMTP
            "Port"          = $MailPortnumber
            "From"          = [MimeKit.MailboxAddress]$MailFrom
            "RecipientList" = [MimeKit.InternetAddressList]@()
            "Subject"       = "WindowsUpdate $MailCustomer $((Get-Date).ToString('MMMM yyyy'))"
            "HTMLBody"      = $result | ConvertTo-Html
        }
        Send-EmailWithRetry -Parameters $EmailParams
    } else {
        Write-Log -LogString "Not the scheduled report day. Exiting script." -LogLevel INFO
    }
} catch {
    Write-Log -LogString "An unexpected error occurred: $_" -LogLevel ERROR
    throw
} finally {
    Write-Log -LogString "Script execution completed." -LogLevel INFO
}
