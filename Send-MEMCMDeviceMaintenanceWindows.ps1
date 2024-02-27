<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Generate htmlpage with Devices and Maintenance Windows
.DESCRIPTION
   Script to be run as schedule task on siteserver. 

   The script generate a html-page and if you use the send-mailkitmessage it will send a mail
   to a group of administrators with info about the Maintenace Windows for a devices in a 
   collection.
.EXAMPLE
   Send-DeviceMaintenanceWindows.ps1

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
$scriptversion = '1.0'
$scriptname = $MyInvocation.MyCommand.Name

$siteserver = 'vntsql0299'
$dbserver = 'VNTSQL0310'
$DaysAfterPatchTuesdayToReport = '-6'
#$DaysAfterPatchTuesdayToReport = '9'
$DisableReport = ""

$filedate = get-date -Format yyyMMdd
$HTMLFileSavePath = "G:\Scripts\Outfiles\KVV_MW_$filedate.HTML"
$CSVFileSavePath = "G:\Scripts\Outfiles\KVV_MW_$filedate.csv"
$SMTP = 'smtp.kvv.se'
$MailFrom = 'no-reply@kvv.se'
$MailTo1 = 'christian.damberg@kriminalvarden.se'
#$MailTo2 = 'Joakim.Stenqvist@kriminalvarden.se'
#$mailto3 = 'Julia.Hultkvist@kriminalvarden.se'
#$mailto4 = 'Christian.Brask@kriminalvarden.se'
#$mailto5 = 'lars.garlin@kriminalvarden.se'
#$MailTo6 = 'sockv@kriminalvarden.se'
#$mailto7 = 'Tim.Gustavsson@kriminalvarden.se'
$MailPortnumber = '25'
$MailCustomer = 'Kriminalvården - IT'
$collectionidToCheck = 'KV1000B0'

$Logfile = "G:\Scripts\Logfiles\WindowsUpdateScript.log"
$logfile_MW = "G:\Scripts\Logfiles\MW_Script_$filedate.log"
function Write-Log
{
Param ([string]$LogString)
$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $LogFile -value $LogMessage
}

function Write-MWLog
{
Param ([string]$LogString)
$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $logfile_MW -value $LogMessage
}

function Rotate-log 
{
    $target = Get-ChildItem $Logfile -Filter "windows*.log"
    $threshold = 30
    $datetime = Get-Date -uformat "%Y-%m-%d-%H%M"
    $target | ForEach-Object {
    if ($_.Length -ge $threshold) { 
        Write-Host "file named $($_.name) is bigger than $threshold KB"
        $newname = "$($_.BaseName)_${datetime}.log_old"
        Rename-Item $_.fullname $newname
        Write-Host "Done rotating file" 
    }
    else{
         Write-Host "file named $($_.name) is not bigger than $threshold KB"
    }
    Write-Host " "
}
}



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
	#Install-Module send-mailkitmessage -ErrorAction SilentlyContinue
	Import-Module send-mailkitmessage
}


if (-not (Get-Module -name PSWriteHTML))
{
	#Install-Module PSWriteHTML -ErrorAction SilentlyContinue
	Import-Module PSWriteHTML
}

if (-not (Get-Module -name PatchManagementSupportTools))
{
	#Install-Module PatchManagementSupportTools -ErrorAction SilentlyContinue
	Import-Module PatchManagementSupportTools
}

#endregion

function Get-CMSiteCode
{
	$CMSiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
	return $CMSiteCode
}


Get-CMModule
#Write-Log -LogString "Send-DeviceMaintenanceWindows - CMmodule imported"
$sitecode = get-cmsitecode
#Write-Log -LogString "Send-DeviceMaintenanceWindows - $sitecode extracted"
$SetSiteCode = $sitecode + ":"
Set-Location $SetSiteCode
#Write-Log -LogString "Send-DeviceMaintenanceWindows - set location to $SetSiteCode"

<#
	===========================================================================		
	Date-section

	DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!DON´T EDIT!!!
	===========================================================================
#>

$todayDefault = Get-Date
$todayshort = $todayDefault.ToShortDateString()
$thismonth = $todaydefault.Month
$nextmonth = $todaydefault.Month + 1
$patchtuesdayThisMonth = Get-PatchTuesday -Month $thismonth -Year $todayDefault.Year
$patchtuesdayNextMonth = Get-PatchTuesday -Month $nextmonth -Year $todayDefault.Year
$ReportdayCompare = ($patchtuesdayThisMonth.AddDays($DaysAfterPatchTuesdayToReport)).tostring("yyyy-MM-dd")

$nextyear = $todayDefault.Year + 1
If ($nextmonth = '13')
{
    $nextmonth = '1'
    
}
$checkdatestart = $patchtuesdayThisMonth.ToShortDateString()


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
$counter = 0

#check if script should run or not

if($todayDefault.Month -in $DisableReport)

{

	#write-host "date not equal"
	Write-Log -LogString "Send-DeviceMaintenanceWindows - This month is skipped"
	#write-host -ForegroundColor Green "This month the updates will not be installed"
	
	set-location $PSScriptRoot
	
	exit

}


#Region Script part 1 collect info from selected collection and check devices membership in Collections with Maintenance Windows

if ($todayshort -eq $ReportdayCompare)
{
	# Array to collect data in
	$ResultColl = @()
	$ResultMissing = @()
	# Devices

    $devices = Get-CMCollectionMember -CollectionId $collectionidToCheck
	Write-Log -LogString "$scriptname - Date is correct, will run script"
	# For the progressbar
	$complete = 0
	
	# Loop for each device
	foreach ($device in $devices)
	{
		$counter++
		Write-Progress -Activity 'Processing computer' -CurrentOperation $device.Name -PercentComplete (($counter / $devices.count) * 100)
		Start-Sleep -Milliseconds 100
        
        $Devicename = $device.name

        $Computertotal = $devices.Count
		Write-MWLog -LogString "$scriptname - Processing computer...$Devicename -  $counter of $Computertotal"
		# Get all Collections for Device

        try
            {
    	        $collectionids = Get-CMClientDeviceCollectionMembership -ComputerName $Devicename -SiteServer $siteserver -SiteCode $sitecode
	        }
        catch
            {
                Write-Log -LogString "Can´t get any collectionid for $Devicename $_"
            }
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
						if ($mw.StartTime -gt $patchtuesdayThisMonth -and $mw.StartTime -lt $patchtuesdayNextMonth)
						{
							$computername = $device.Name

                            $query = "SELECT applikation FROM tblinmatning WHERE skrotad=0 AND servernamn='$Computername'"
                            Write-MWLog -LogString "Query against extra db run $computername"
            
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
	
	
	$ResultColl | Export-Csv -Path $CSVFileSavePath -Encoding UTF8
	Write-Log -LogString "$scriptname - File $CSVFileSavePath created"
    $ResultColl.Count
}

else
{
	Write-Log -LogString "$scriptname - Date not equal patchtuesday $checkdatestart and its now $todayshort. This report will run $ReportdayCompare"
	Write-Log -LogString "$scriptname - Script exit!"
    
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
		
		New-HTMLTable -DataTable $ResultColl -PagingLength 35 -Style compact
		
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
med fönster mellan $patchtuesdayThisMonth och $patchtuesdayNextMonth<br>
<p>Se bifogad fil. Kom ihåg att kopiera planen till<br>
\\kvv.se\dokument\ProjektKVS\IT_Enheten\ITIL Processer\Change\Winpatchar
</p>
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
$RecipientList.Add([MimeKit.InternetAddress]$MailTo1)
#$RecipientList.Add([MimeKit.InternetAddress]$MailTo2)
#$RecipientList.Add([MimeKit.InternetAddress]$mailto3)
#$RecipientList.Add([MimeKit.InternetAddress]$MailTo4)
#$RecipientList.Add([MimeKit.InternetAddress]$mailto5)
#$RecipientList.add([MimeKit.InternetAddress]$mailto6)
#$RecipientList.add([MimeKit.InternetAddress]$mailto7)
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
Write-Log -LogString "$scriptname - Mail on it´s way to $RecipientList"
set-location $PSScriptRoot
Write-Log -LogString "$scriptname - Script exit!"
Write-MWLog -LogString "Script done!"
#endregion
