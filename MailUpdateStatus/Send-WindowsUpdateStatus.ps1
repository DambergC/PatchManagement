<#
	.SYNOPSIS
		Updatestatus for Windows update thru MECM
	
	.DESCRIPTION
		The script create a a report on windows update for one or more deployments of updates specified in a xml file and sends a email to named recipients.
		The script can run manually or scheduled on siteserver.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2024
		Created on:   	10/16/2023 3:34 PM
		Updated on:     03/25/2024 4:00 PM
		Created by:   	Christian Damberg
		Organization:	Telia Cygate AB
		Filename:	    Send-WindowsUpdateStatus.ps1
		===========================================================================
#>

try {
    $xml = [System.Xml.XmlDocument]::new()
    $xml.Load(".\ScriptConfig.xml")
} catch {
    Write-Log -LogString "Error loading XML file: $_"
    exit 1
}

$Logfilepath = $xml.Configuration.Logfile.Path
$logfilename = $xml.Configuration.Logfile.Name
$Logfile = $Logfilepath + $logfilename
$Logfilethreshold = $xml.Configuration.Logfile.Logfilethreshold


$XMLREAD = $xml.Configuration.Runscript.Job
$scriptname = $MyInvocation.MyCommand.Name
$DisableReport = $xml.Configuration.DisableReportMonth | ForEach-Object {$_.DisableReportMonth.Number}
$siteserver = $xml.Configuration.SiteServer
$filedate = get-date -Format yyyMMdd
$SMTP = $xml.Configuration.MailSMTP
$MailFrom = $xml.Configuration.Mailfrom
$MailPortnumber = $xml.Configuration.MailPort
$MailCustomer = $xml.Configuration.MailCustomer

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
                Write-Log -LogString "Rotated log file: $($file.Name)"
            } catch {
                Write-Log -LogString "Failed to rotate log file $($file.Name): $_"
            }
        } else {
            Write-Log -LogString "Log file $($file.Name) does not need rotation."
        }
    }
}

Rotate-Log

Function Write-Log {
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

Function Get-CMSiteCode
    {
        $CMSiteCode = Get-CimInstance -Namespace "root\SMS" -ClassName SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
    	return $CMSiteCode
    }

Function Ensure-Module {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -Name $ModuleName -ListAvailable)) {
        try {
            Install-Module -Name $ModuleName -Force -ErrorAction Stop
            Import-Module -Name $ModuleName
        } catch {
            Write-Log -LogString "Failed to install or import module $ModuleName: $_"
        }
    } else {
        Import-Module -Name $ModuleName
    }
}

Ensure-Module -ModuleName "send-mailkitmessage"
Ensure-Module -ModuleName "PSWriteHTML"

# PatchManagementSupportTools - Created by Christian Damberg, Cygate
# https://github.com/DambergC/PatchManagement/tree/main/PatchManagementSupportTools

if (-not (Get-Module -name PatchManagementSupportTools))
    {
        #Install-Module PatchManagementSupportTools -ErrorAction SilentlyContinue
    	Import-Module PatchManagementSupportTools
    }

Get-CMModule
$sitecode = get-cmsitecode
$SetSiteCode = $sitecode + ":"
Set-Location $SetSiteCode

<#
	===========================================================================		
	Date-section
	===========================================================================
#>
Write-Log -LogString "============================== $scriptname - Script START ====================================="

$ParametersNode = $xml.configuration.Runscript.SelectNodes('//Job')
$todayDefault = Get-Date
$todayshort = $todayDefault.ToShortDateString()
$thismonth = $todaydefault.Month
$nextmonth = $todaydefault.Month + 1
$patchtuesdayThisMonth = Get-PatchTuesday -Month $thismonth -Year $todayDefault.Year
$patchtuesdayNextMonth = Get-PatchTuesday -Month $nextmonth -Year $todayDefault.Year
$checkdatestart = $patchtuesdayThisMonth.ToShortDateString()
$nextyear = $todayDefault.Year + 1

If ($nextmonth = '13')
{
    $nextmonth = '1'
}
If ($nextmonth = '13')
{
    $nextyear = ((get-date).Year) +1
    $checkdateend = Get-PatchTuesday -Month '1' -Year  $nextyear
}
else
{
    $checkdateend = $patchtuesdayNextMonth.ToShortDateString()
}

$TitleDate = get-date -DisplayHint Date

#check if script should run or not
if($todayDefault.Month -in $DisableReport)
{
    Write-Log -LogString "======================== $scriptname - This month is skipped =============================="
	Write-Log -LogString "============================ $scriptname - script exit =================================="
	set-location $PSScriptRoot
	exit
}



# Loop over selected nodes
foreach($Node in $ParametersNode){
$HTMLfilepath = $xml.Configuration.HTMLfilePath 

$DescToFile = $node.Description
$filename = $DescToFile + "_UpdateStatus_" + $filedate + ".HTML"
$HTMLFileSavePath = "$HTMLfilepath\$filename" # Path to attached file

$deploymentIDtoCheck = $node.DeploymentID
$DaysAfterPatchTuesdayToReport = $node.Offsetdays
$ReportdayCompare = ($patchtuesdayThisMonth.AddDays($DaysAfterPatchTuesdayToReport)).tostring("yyyy-MM-dd")
$GroupToRun = $node.Description

Write-Log -LogString "========================================== $GroupToRun ================================================="

<#
	===========================================================================		
	Collect data from Deployment
	===========================================================================
#>

$ResultColl = @()

if ($todayshort -eq $ReportdayCompare)

{
    Write-Log -LogString "$scriptname - Script starts to collect data"
	$UpdateStatus = Get-SCCMSoftwareUpdateStatus -DeploymentID $deploymentIDtoCheck
	
	foreach ($US in $UpdateStatus)
	
        {
        $object = New-Object -TypeName PSObject
        		$object | Add-Member -MemberType NoteProperty -Name 'Server' -Value $us.DeviceName
        		$object | Add-Member -MemberType NoteProperty -Name 'Collection' -Value $us.CollectionName
        		$object | Add-Member -MemberType NoteProperty -Name 'Status' -Value $us.Status
        		$object | Add-Member -MemberType NoteProperty -Name 'StatusTid' -Value $us.StatusTime
        		
        		$resultColl += $object
        }
}

else

{
    Write-Log -LogString "$scriptname - Date not equal patchtuesday $checkdatestart and its now $todayshort. This report will run $ReportdayCompare"
}

if ($ResultColl -gt $null)

    {
    
    # Create vaules to the report
    
    $errorUnknownvalue = ($UpdateStatus | Where-Object { ($_.status -eq 'error' -or $_.status -eq 'unknown' -or $_.status -eq 'Inprogress') }).count
    $successvalue = ($UpdateStatus | Where-Object { ($_.status -eq 'success') }).count
    $ToCheck = $UpdateStatus | Where-Object { ($_.status -eq 'error' -or $_.status -eq 'unknown' -or $_.status -eq 'Inprogress') } | ConvertTo-Html
    $colletionname = $UpdateStatus.collectionname | Select-Object -First 1
    
    <#
	===========================================================================		
	HTML-time, create the report
	===========================================================================
    #>
    
    New-HTML -TitleText "Uppdatering Status - $MailCustomer" -FilePath $HTMLFileSavePath -ShowHTML -Online {
	
	New-HTMLHeader {
		New-HTMLSection -Invisible {
			New-HTMLPanel -Invisible {
				New-HTMLText -Text "$MailCustomer - UpdateStatus" -FontSize 35 -Color Darkblue -FontFamily Arial -Alignment center
				New-HTMLHorizontalLine
			}
		}
	}
	New-HTMLSection -Invisible -Title "UpdateStatus $filedate"{
		
		New-HTMLPanel {
			New-HTMLChart {
				New-ChartLegend -LegendPosition bottom -HorizontalAlign right -Color red, darkgreen -DisableOnItemClickToggleDataSeries -DisableOnItemHoverHighlightDataSeries
				New-ChartAxisY -LabelMaxWidth 100 -LabelAlign left -Show -LabelFontColor red, darkgreen -TitleText 'Status' -TitleColor Red
				New-ChartBarOptions -Distributed
				New-ChartBar -Name 'Needs attention' -Value $errorUnknownvalue
                New-ChartBar -name 'Success' -Value $successvalue
			} -Title 'Resultat av patchning' -TitleAlignment center -SubTitle $colletionname -SubTitleAlignment center -SubTitleFontSize 20 -TitleColor Darkblue
		}
	}
	
	New-HTMLSection -Invisible -Title "UpdateStatus $filedate"{
		
		New-HTMLTable -DataTable $resultColl -PagingLength 25 -Style compact
	}
	New-HTMLFooter {
		New-HTMLSection -Invisible {
			New-HTMLPanel -Invisible {
				New-HTMLHorizontalLine
				New-HTMLText -Text "This report was created $today" -FontSize 20 -Color Darkblue -FontFamily Arial -Alignment center -FontStyle italic
			}
		}
	}
    }
    
    $Body = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Windowsupdate Status - $MailCustomer</title>
<style>
    th {
        font-family: Arial, Helvetica, sans-serif;
        color: White;
        font-size: 12px;
        border: 1px solid black;
        padding: 3px;
        background-color: Black;
    } 
    p {
        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 12px;
    } 
    ol {
        font-family: Arial, Helvetica, sans-serif;
        list-style-type: square;
        color: black;
        font-size: 12px;
    }
	    H1 {
        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 18px;
            }
    tr {
        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 11px;
        vertical-align: text-top;
    } 

    body {
        background-color: lightgray;
      }
      table {
        border: 1px solid black;
        border-collapse: collapse;
      }

      td {
        border: 1px solid black;
        padding: 5px;
        background-color: #E0F3F7;
      }

</style>
</head>

<body>
    <IMG
SRC="data:image/jpg;base64,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
"ALT="logo">
	<p><h1>WindowsUpdate Status</h1></p> 
	<p>Attached html-file contains status for servers:</p>
    <p> Collection: $colletionname </p>
    <p> Need Attention: $errorUnknownvalue </p>
    <p> Success: $successvalue </p>
    <p> $ToCheck </p>

<hr>
</p> 
	<p>Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
</body>
</html>

"@
    
    <#
	===========================================================================		
	Mailsettings
	===========================================================================
    #>
    
    #use secure connection if available ([bool], optional)
    $UseSecureConnectionIfAvailable = $false
    
    #authentication ([System.Management.Automation.PSCredential], optional)
    $Credential = [System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))
    
    #SMTP server ([string], required)
    $SMTPServer = $SMTP
    
    #port ([int], required)
    $Port = $MailPortnumber
    
    #sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
    $From = [MimeKit.MailboxAddress]$MailFrom
    
    #recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
    $RecipientList = [MimeKit.InternetAddressList]::new()
    
    $recipientlistXML = $xml.Configuration.Recipients | ForEach-Object {$_.Recipients.Email}
    
    foreach ($Recipient in $recipientlistXML)
    
        {
            $RecipientList.Add([MimeKit.InternetAddress]$Recipient)
        }
    
    #cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
    #$CCList=[MimeKit.InternetAddressList]::new()
    #$CCList.Add([MimeKit.InternetAddress]$EmailToCC)
    
    #bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
    #$BCCList = [MimeKit.InternetAddressList]::new()
    #$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")
    
    #subject ([string], required)
    $Subject = [String]"Serverpatchning $MailCustomer $monthname $year - $colletionname"
    
    #text body ([string], optional)
    #$TextBody=[string]"TextBody"
    
    #HTML body ([string], optional)
    $HTMLBody = [String]$Body
    
    #attachment list ([System.Collections.Generic.List[string]], optional)
    $AttachmentList = [System.Collections.Generic.List`1[[System.String, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089]]]::new()
    $AttachmentList.Add("$HTMLFileSavePath")
    #$AttachmentList.Add("$CSVFileSavePath")
    
    # Mailparameters
    $Parameters = @{
    	"UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable
    	#"Credential"=$Credential
    	"SMTPServer"					 = $SMTPServer
    	"Port"						     = $Port
    	"From"						     = $From
    	"RecipientList"				     = $RecipientList
    	#"CCList"=$CCList
    	#"BCCList"=$BCCList
    	"Subject"					     = $Subject
    	#"TextBody"=$TextBody
    	"HTMLBody"					     = $HTMLBody
    	"AttachmentList"				 = $AttachmentList
    }
    Write-Log -LogString "========================================== $GroupToRun ================================================="
    Write-Log -LogString "=========================================== Sending mail ==============================================="
    Write-Log -LogString "$scriptname - Script send mail to $RecipientList"
    Send-MailKitMessage @Parameters
    
    
    
    }


    else

{
    Write-Log -LogString "========================================== $GroupToRun ================================================="
    Write-Log -LogString "================================= Scriptrunday not equal with xmlfile for deployment ===============================" 
    
 
}


}

Write-Log -LogString "================================= $scriptname - script exit ===================================="



set-location $PSScriptRoot
