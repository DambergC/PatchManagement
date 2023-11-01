<#
	.SYNOPSIS
		Collects data for a specific updategroup
	
	.DESCRIPTION
		Summarize about the updategroup status and present it in an html-document sent by email.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2023 v5.8.232
		Created on:   	10/16/2023 3:34 PM
		Created by:   	damberg
		Organization:	Telia Cygate AB
		Filename:		Send-MEMCMUpdateStatus.ps1
		===========================================================================
#>

<#
	===========================================================================
	Values needed to be updated before running the script
	===========================================================================
#>


$deploymentIDtoCheck = '16777362'
$DaysAfterPatchTuesdayToReport = '6'
$siteserver = 'vntsql0299'
$filedate = get-date -Format yyyMMdd
$HTMLFileSavePath = "c:\temp\$sitecode_UpdateStatus_$filedate.HTML"
$CSVFileSavePath = "c:\temp\$sitecodeUpdateStatus_$filedate.csv"

$SMTP = 'smtp.kvv.se'
$MailFrom = 'no-reply@kvv.se'
$MailTo = 'christian.damberg@kriminalvarden.se'
$MailPortnumber = '25'
$MailCustomer = 'Kriminalvården - IT'


<#
	===========================================================================
	Powershell modules needed in the script
	===========================================================================

	Send-MailkitMessage - https://github.com/austineric/Send-MailKitMessage

	pswritehtml - https://github.com/EvotecIT/PSWriteHTML

	PatchManagementSupportTools - Created by Christian Damberg, Cygate
	https://github.com/DambergC/PatchManagement/tree/main/PatchManagementSupportTools

	DON´T EDIT!!!
#>

#region modules

if (-not (Get-Module -name send-mailkitmessage))
{
	Install-Module send-mailkitmessage -ErrorAction SilentlyContinue
	Import-Module send-mailkitmessage
	write-host -ForegroundColor Green 'Send-Mailkitmessage imported'
}

else
{
	
	write-host -ForegroundColor Green 'Send-Mailkitmessage already imported and installed!'
}


if (-not (Get-Module -name PSWriteHTML))
{
	Install-Module PSWriteHTML -ErrorAction SilentlyContinue
	Import-Module PSWriteHTML
	write-host -ForegroundColor Green 'PSWriteHTML imported'
}

else
{
	
	write-host -ForegroundColor Green 'PSWriteHTML already imported and installed!'
}

if (-not (Get-Module -name PatchManagementSupportTools))
{
	Install-Module PatchManagementSupportTools -ErrorAction SilentlyContinue
	Import-Module PatchManagementSupportTools
	write-host -ForegroundColor Green 'PatchManagementSupportTools imported'
}

else
{
	
	write-host -ForegroundColor Green 'PatchManagementSupportTools already imported and installed!'
}



#endregion

<#
	===========================================================================
	Parameters needed in the script
	===========================================================================

	Here youy may edit what you need to make the script working in your environment.
#>


function Get-CMSiteCode
{
	$CMSiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
	return $CMSiteCode
}


Get-CMModule
$sitecode = get-cmsitecode
$SetSiteCode = $sitecode + ":"
Set-Location $SetSiteCode

$ResultColl = @()

<#
	===========================================================================		
	Date-section
	===========================================================================
#>

$todayDefault = Get-Date
$todayCompare = (get-date).ToString("yyyy-MM-dd")
$patchdayDefault = Get-PatchTuesday -Month $todayDefault.Month -Year $todayDefault.Year
$patchdayCompare = (Get-PatchTuesday -Month $todayDefault.Month -Year $todayDefault.Year).tostring("yyyy-MM-dd")

# Compare and see if the report should be run or not...
$ReportdayCompare = ($patchdayDefault.AddDays($DaysAfterPatchTuesdayToReport)).tostring("yyyy-MM-dd")


<#
	===========================================================================		
	Collect data from Deployment
	===========================================================================
#>


if ($todayCompare -eq $ReportdayCompare)
{
	
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
	
	write-host "date not equal"
	
	write-host -ForegroundColor Green "Patch tuesday is $patchdayCompare and Today it is $todayCompare and rundate for the report is $ReportdayCompare"
	
	exit
	
}

# Create vaules to the report

$errorvalue = ($UpdateStatus | Where-Object { ($_.status -eq 'error') }).count
$successvalue = ($UpdateStatus | Where-Object { ($_.status -eq 'success') }).count
$colletionname = $UpdateStatus.collectionname | Select-Object -First 1

<#
	===========================================================================		
	HTML-time, create the report
	===========================================================================
#>

New-HTML -TitleText "Uppdatering Status - Kriminalvården" -FilePath $HTMLFileSavePath -ShowHTML -Online {
	
	New-HTMLHeader {
		New-HTMLSection -Invisible {
			New-HTMLPanel -Invisible {
				New-HTMLText -Text "Kriminalvården - UpdateStatus" -FontSize 35 -Color Darkblue -FontFamily Arial -Alignment center
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
				New-ChartBar -Name 'Error' -Value $errorvalue
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
				New-HTMLText -Text "Denna lista skapades $today" -FontSize 20 -Color Darkblue -FontFamily Arial -Alignment center -FontStyle italic
			}
			
		}
	}
	
}

$Body = @"

<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Uppdatering Status - Kriminalvården</title>
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
	<p><h1>Uppdatering Status</h1></p> 
	<p>Bifogade filer innehåller Status för aktuell update Grupp.<br><br>
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
$RecipientList.Add([MimeKit.InternetAddress]$MailTo)


#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$CCList=[MimeKit.InternetAddressList]::new()
#$CCList.Add([MimeKit.InternetAddress]$EmailToCC)



#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList = [MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")


#subject ([string], required)
$Subject = [string]"Serverpatchning $MailCustomer $monthname $year"

#text body ([string], optional)
#$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody = [string]$Body

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList = [System.Collections.Generic.List[string]]::new()
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



Send-MailKitMessage @Parameters

set-location $PSScriptRoot




