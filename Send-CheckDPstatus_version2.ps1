<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
    Check Distribution Points (DP) and send an email report with their statuses.
.DESCRIPTION
    This script checks the status of DPs, logs the results, and sends an email report with detailed information.
    It uses modularized functions, enhanced error handling, and improved logging for better maintainability.

.DISCLAIMER
    All scripts and PowerShell references are offered AS IS with no warranty. 
    Test scripts in a test environment before using them in production.
-------------------------------------------------------------------------------------------------------------------------
#>

# Configuration Variables
$Config = @{
    ScriptName         = $MyInvocation.MyCommand.Name
    SiteServer         = 'vntsql0299'
    ProviderMachine    = "vntsql0299.kvv.se"
    DPMaintGroup       = 'Maintenance'
    DPProdGroup        = 'VS DPs'

    # Mail Settings
    FileDate           = Get-Date -Format yyyMMdd
    MailFrom           = 'no-reply@kvv.se'
    MailRecipients     = @('christian.damberg@kriminalvarden.se', 'Joakim.Stenqvist@kriminalvarden.se', 'Julia.Hultkvist@kriminalvarden.se', 'Tim.Gustavsson@kriminalvarden.se')
    MailSMTP           = 'smtp.kvv.se'
    MailPort           = 25
    MailCustomer       = 'Kriminalvården'

    # Logfile
    LogFilePath        = "G:\Scripts\Logfiles\DPLogfile.log"
}

# Import Required Modules
function Import-RequiredModules {
    Write-Log -LogString "Importing Required Modules..."
    @('ConfigurationManager', 'send-mailkitmessage', 'PSWriteHTML') | ForEach-Object {
        if (-not (Get-Module -Name $_)) {
            Import-Module $_ -ErrorAction Stop
            Write-Log -LogString "Module $_ imported successfully."
        }
    }
}

# Logging Function
function Write-Log {
    param ([string]$LogString)
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-Content -Path $Config.LogFilePath -Value $LogMessage
}

# Get Site Code
function Get-CMSiteCode {
    Try {
        $CMSiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $Config.SiteServer | Select-Object -ExpandProperty SiteCode
        return $CMSiteCode
    } Catch {
        Write-Log -LogString "ERROR: Failed to retrieve site code. Message: $_"
        throw
    }
}

# Check Distribution Points
function Check-DPs {
    Write-Log -LogString "Fetching DP information..."
    $DPs = Get-CMDistributionPoint -SiteCode $SiteCode | Select-Object NetworkOSPath | Sort-Object -Property NetworkOSPath

    $SucceededDPs = @()
    $FailedDPs = @()

    foreach ($DP in $DPs) {
        $DPName = ($DP.NetworkOSPath -replace "\\", "").ToUpper()

        if (Test-Connection -ComputerName $DPName -Count 4 -Quiet) {
            Write-Log -LogString "$DPName is online."
            $SucceededDPs += $DPName
        } else {
            Write-Log -LogString "$DPName is not online."
            $FailedDPs += $DPName
        }
    }

    return @{
        Succeeded = $SucceededDPs
        Failed    = $FailedDPs
    }
}

# Update DP Groups and Maintenance Mode
function Update-DPStatus {
    param (
        [array]$DPs,
        [bool]$EnableMaintenanceMode,
        [string]$AddToGroup,
        [string]$RemoveFromGroup
    )

    foreach ($DP in $DPs) {
        try {
            Set-CMDistributionPoint -SiteSystemServerName $DP -EnableMaintenanceMode $EnableMaintenanceMode -Force
            Add-CMDistributionPointToGroup -DistributionPointName $DP -DistributionPointGroupName $AddToGroup
            Remove-CMDistributionPointFromGroup -DistributionPointName $DP -DistributionPointGroupName $RemoveFromGroup -Force
            Write-Log -LogString "$DP updated: MaintenanceMode=$EnableMaintenanceMode, Added to $AddToGroup, Removed from $RemoveFromGroup."
        } Catch {
            Write-Log -LogString "ERROR: Failed to update $DP. Message: $_"
        }
    }
}

# Send Email Report
function Send-EmailReport {
    param (
        [string]$HTMLBody
    )
    $Parameters = @{
        SMTPServer                   = $Config.MailSMTP
        Port                         = $Config.MailPort
        From                         = $Config.MailFrom
        RecipientList                = [MimeKit.InternetAddressList]::new() + $Config.MailRecipients
        Subject                      = "Distribution Point Status Report - $(Get-Date -Format yyyy-MM-dd)"
        HTMLBody                     = $HTMLBody
        UseSecureConnectionIfAvailable = $false
    }

    Try {
        Send-MailKitMessage @Parameters
        Write-Log -LogString "Email sent successfully."
    } Catch {
        Write-Log -LogString "ERROR: Failed to send email. Message: $_"
    }
}

# Main Script Execution
try {
    Write-Log -LogString "Script started."
    Import-RequiredModules

    $SiteCode = Get-CMSiteCode
    Write-Log -LogString "Site code retrieved: $SiteCode"

    $DPStatus = Check-DPs
    Write-Log -LogString "DPs Checked. Online: $($DPStatus.Succeeded.Count), Offline: $($DPStatus.Failed.Count)"

    # Update Offline DPs
    if ($DPStatus.Failed.Count -gt 0) {
        Update-DPStatus -DPs $DPStatus.Failed -EnableMaintenanceMode $true -AddToGroup $Config.DPMaintGroup -RemoveFromGroup $Config.DPProdGroup
    }

    # Update Online DPs
    if ($DPStatus.Succeeded.Count -gt 0) {
        Update-DPStatus -DPs $DPStatus.Succeeded -EnableMaintenanceMode $false -AddToGroup $Config.DPProdGroup -RemoveFromGroup $Config.DPMaintGroup
    }

    # Generate HTML Report
    $HTMLBody = "<html><body><h1>DP Status Report</h1><p>Online: $($DPStatus.Succeeded.Count), Offline: $($DPStatus.Failed.Count)</p></body></html>"

    # Send Email
    Send-EmailReport -HTMLBody $HTMLBody

    Write-Log -LogString "Script completed successfully."
} Catch {
    Write-Log -LogString "ERROR: Script failed. Message: $_"
} finally {
    Write-Log -LogString "Script end."
}
