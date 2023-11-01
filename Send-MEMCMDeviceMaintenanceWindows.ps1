<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Generate htmlpage with Devices and Maintenance Windows
.DESCRIPTION
   Script to be run as schedule task on siteserver. It's recommended to be use my script to
   Generate scheduleTask based on offset from patchTuesday.

   The script generate a html-page and if you use the send-mailkitmessage it will send a mail
   to a group of administrators with info about the Maintenace Windows for a devices in a 
   collection.
.EXAMPLE
   Show-DeviceMaintenanceWindows.ps1

.DISCLAIMER
All scripts and other Powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
-------------------------------------------------------------------------------------------------------------------------
#>

<#
	===========================================================================
	Values needed to be updated before running the script
	===========================================================================
#>

$siteserver = 'cm01'
$dbserver = '<SQLserver_for_extra_info>'
$DaysAfterPatchTuesdayToReport = '6'

$HTMLFileSavePath = "c:\temp\KVV_MW_$filedate.HTML"
$CSVFileSavePath = "c:\temp\KVV_MW_$filedate.csv"
$SMTP = '<SMTP-server>'
$MailFrom = '<no-replyaddress>'
$MailTo = '<Mailto>'
$MailPortnumber = '25'
$MailCustomer = '<Customername>'
$collectionidToCheck = '<Collectionid>'





<#
	===========================================================================
	Powershell modules needed in the script
	===========================================================================

	Send-MailkitMessage - https://github.com/austineric/Send-MailKitMessage

	pswritehtml - https://github.com/EvotecIT/PSWriteHTML

	PatchManagementSupportTools - Created by Christian Damberg, Cygate
	https://github.com/DambergC/PatchManagement/tree/main/PatchManagementSupportTools

	DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!
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

function Get-CMSiteCode
{
	$CMSiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
	return $CMSiteCode
}


Get-CMModule -Verbose
$sitecode = get-cmsitecode

$SetSiteCode = $sitecode + ":"
Set-Location $SetSiteCode

<#
	===========================================================================		
	Date-section

	DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!
	===========================================================================
#>

$todayDefault = Get-Date
$todayCompare = (get-date).ToString("yyyy-MM-dd")
$patchdayDefault = Get-PatchTuesday -Month $todayDefault.Month -Year $todayDefault.Year
$patchdayCompare = (Get-PatchTuesday -Month $todayDefault.Month -Year $todayDefault.Year).tostring("yyyy-MM-dd")

# Compare and see if the report should be run or not...
$ReportdayCompare = ($patchdayDefault.AddDays($DaysAfterPatchTuesdayToReport)).tostring("yyyy-MM-dd")


# Date and mail section
$todaydefault = Get-Date
$nextmonth = $todaydefault.Month + 1

$checkdatestart = Get-PatchTuesday -Month $todaydefault.Month -Year $todaydefault.Year
$checkdateend = Get-PatchTuesday -Month $nextmonth -Year $todaydefault.Year
$filedate = get-date -Format yyyMMdd
$TitleDate = get-date -DisplayHint Date
$counter = 0




#Region Script part 1 collect info from selected collection and check devices membership in Collections with Maintenance Windows

if ($todayCompare -eq $ReportdayCompare)
{
	# Array to collect data in
	$ResultColl = @()
	$ResultMissing = @()
	# Devices
	$devices = Get-CMCollectionMember -CollectionId $collectionidToCheck
	
	
	# For the progressbar
	$complete = 0
	
	
	# Loop for each device
	foreach ($device in $devices)
	{
		$counter++
		Write-Progress -Activity 'Processing computer' -CurrentOperation $device.Name -PercentComplete (($counter / $devices.count) * 100)
		Start-Sleep -Milliseconds 100
		
		# Get all Collections for Device
		$collectionids = Get-CMClientDeviceCollectionMembership -ComputerName $device.name
		
		# Check every Collection for Maintenance windows
		foreach ($collectionid in $collectionids)
		{
			
			# Only include Collections with Maintenance Windows
			if ($collectionid.ServiceWindowsCount -gt 0)
			{
				$MWs = Get-CMMaintenanceWindow -CollectionId $collectionid.CollectionID
				
				foreach ($mw in $MWs)
				{
					
					if ($mw.RecurrenceType -eq 1)
					{
						# Only show Maintenance Windows waiting to run
						if ($mw.StartTime -gt $checkdatestart -and $mw.StartTime -lt $checkdateend)
						{
							$computername = $device.Name
							$query = "SELECT applikation FROM tblinmatning WHERE skrotad=0 AND servernamn='$Computername'"
							$data = Invoke-Sqlcmd -ServerInstance $dbserver -Database serverlista -Query $query
							$Startdatum = ($mw.StartTime).ToString("yyyy-MM-dd")
							$starttid = ($mw.StartTime).ToString("hh:mm")
							
							$object = New-Object -TypeName PSObject
							$object | Add-Member -MemberType NoteProperty -Name 'Applikation' -Value $data.applikation
							$object | Add-Member -MemberType NoteProperty -Name 'Server' -Value $device.name
							$object | Add-Member -MemberType NoteProperty -Name 'Startdatum' -Value $Startdatum
							$object | Add-Member -MemberType NoteProperty -Name 'Starttid' -Value $starttid
							$object | Add-Member -MemberType NoteProperty -Name 'Varaktighet' -Value $mw.Duration
							$object | Add-Member -MemberType NoteProperty -Name 'Deployment' -Value $collectionid.name
							$resultColl += $object
						}
						
					}
					
					if ($mw.RecurrenceType -eq 3)
					{
						
						$computername = $device.Name
						$query = "SELECT applikation FROM tblinmatning WHERE skrotad=0 AND servernamn='$Computername'"
						$data = Invoke-Sqlcmd -ServerInstance $dbserver -Database serverlista -Query $query
						$Startdatum = ($mw.StartTime).ToString("yyyy-MM-dd")
						$starttid = ($mw.StartTime).ToString("hh:mm")
						
						$object = New-Object -TypeName PSObject
						$object | Add-Member -MemberType NoteProperty -Name 'Applikation' -Value $data.applikation
						$object | Add-Member -MemberType NoteProperty -Name 'Server' -Value $device.name
						$object | Add-Member -MemberType NoteProperty -Name 'Startdatum' -Value $Startdatum
						$object | Add-Member -MemberType NoteProperty -Name 'Starttid' -Value $mw.Name
						$object | Add-Member -MemberType NoteProperty -Name 'Varaktighet' -Value $mw.Duration
						$object | Add-Member -MemberType NoteProperty -Name 'Deployment' -Value $collectionid.name
						$resultColl += $object
					}
					
					
				}
				
				
			}
			
		}
	}
	
	
	$ResultColl | Export-Csv -Path $CSVFileSavePath -Encoding UTF8 -Verbose
	
}

else
{
	
	write-host "date not equal"
	
	write-host -ForegroundColor Green "Patch tuesday is $patchdayCompare and Today it is $todayCompare and rundate for the report is $ReportdayCompare"
	
	set-location $PSScriptRoot
	
	exit
	
	
}


#endregion

#region Script part 2 Create the html-file to be distributed

New-HTML -TitleText "Patchfönster- Kriminalvården" -FilePath $HTMLFileSavePath -ShowHTML -Online {
	
	New-HTMLHeader {
		New-HTMLSection -Invisible {
			New-HTMLPanel -Invisible {
				New-HTMLText -Text "Kriminalvården - Patchfönster" -FontSize 35 -Color Darkblue -FontFamily Arial -Alignment center
				New-HTMLHorizontalLine
			}
		}
	}
	
	New-HTMLSection -Invisible -Title "Maintenance Windows $filedate"{
		
		New-HTMLTable -DataTable $ResultColl -PagingLength 25 -Style compact
		
	}
	
	New-HTMLFooter {
		
		New-HTMLSection -Invisible {
			
			New-HTMLPanel -Invisible {
				New-HTMLHorizontalLine
				New-HTMLText -Text "Denna lista skapades $todaydefault" -FontSize 20 -Color Darkblue -FontFamily Arial -Alignment center -FontStyle italic
			}
			
		}
	}
}

#endregion

#Region CSS and HTML for mail thru Send-MailKitMessage



#endregion

#Region HTML Mail

#Variable needed in html
$collectionname = (Get-CMCollection -id $collectionidToCheck).name


$Body = @"

<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Server Mainenance Windows - Kriminalvården</title>
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
	<p><h1>Server Maintenance Windows - List</h1></p> 
	<p>Bifogad fil innehåller servrar från collection $collectionname.<br><br>
med fönster mellan $checkdatestart och $checkdateend<br>
<hr>
</p> 
	<p>Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>

	
	
	
</body>
</html>
 

"@




#endregion

#Region Mailsettings


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
$AttachmentList.Add("$CSVFileSavePath")

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

#endregion

#Region Send Mail

Send-MailKitMessage @Parameters

set-location $PSScriptRoot


#endregion

