


    Import-Module send-mailkitmessage
    write-host -ForegroundColor Green 'Send-Mailkitmessage imported'
    Import-module PSWriteHTML
    write-host -ForegroundColor Green 'PSWriteHtml imported'

$siteserver = 'vntsql0299'
$filedate = get-date -Format yyyMMdd
$HTMLFileSavePath = "c:\temp\KVV_UpdateStatus_$filedate.HTML"
$CSVFileSavePath = "c:\temp\KVV_UpdateStatus_$filedate.csv"

$SMTP = 'smtp.kvv.se'
$MailFrom = 'no-reply@kvv.se'
$MailTo = 'christian.damberg@kriminalvarden.se'
$MailPortnumber = '25'
$MailCustomer = 'Kriminalvården - IT'

$ResultColl = @()

#region Functions needed in script

function Get-SCCMSoftwareUpdateStatus {
<#
.Synopsis
    This will output the device status for the Software Update Deployments within SCCM.
    For updated help and examples refer to -Online version.
  
 
.DESCRIPTION
    This will output the device status for the Software Update Deployments within SCCM.
    For updated help and examples refer to -Online version.
 
 
.NOTES   
    Name: Get-SCCMSoftwareUpdateStatus
    Author: The Sysadmin Channel
    Version: 1.0
    DateCreated: 2018-Nov-10
    DateUpdated: 2018-Nov-10
 
.LINK
    https://thesysadminchannel.com/get-sccm-software-update-status-powershell -
 
 
.EXAMPLE
    For updated help and examples refer to -Online version.
 
#>
 
    [CmdletBinding()]
 
    param(
        [Parameter()]
        [switch]  $DeploymentIDFromGUI,
 
        [Parameter(Mandatory = $false)]
        [Alias('ID', 'AssignmentID')]
        [string]   $DeploymentID,
         
        [Parameter(Mandatory = $false)]
        [ValidateSet('Success', 'InProgress', 'Error', 'Unknown')]
        [Alias('Filter')]
        [string]  $Status
 
 
    )
 
    BEGIN {
        $Site_Code   = 'KV1'
        $Site_Server = 'vntsql0299'
        $HasErrors   = $False
 
        if ($Status -eq 'Success') {
            $StatusType = 1 
        }
 
        if ($Status -eq 'InProgress') {
            $StatusType = 2
        }
 
        if ($Status -eq 'Unknown') {
            $StatusType = 4
        }
 
        if ($Status -eq 'Error') {
            $StatusType = 5
        }
 
    }
 
    PROCESS {
        try {
            if ($DeploymentID -and $DeploymentIDFromGUI) {
                Write-Error "Select the DeploymentIDFromGUI or DeploymentID Parameter. Not Both"
                $HasErrors   = $True
                throw
            }
 
            if ($DeploymentIDFromGUI) {
                $ShellLocation = Get-Location
                Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
                 
                #Checking to see if module has been imported. If not abort.
                if (Get-Module ConfigurationManager) {
                        Set-Location "$($Site_Code):\"
                        $DeploymentID = Get-CMSoftwareUpdateDeployment | select AssignmentID, AssignmentName | Out-GridView -OutputMode Single -Title "Select a Deployment and Click OK" | Select -ExpandProperty AssignmentID
                        Set-Location $ShellLocation
                    } else {
                        Write-Error "The SCCM Module wasn't imported successfully. Aborting."
                        $HasErrors   = $True
                        throw
                }
            }
 
            if ($DeploymentID) {
                    $DeploymentNameWithID = Get-WMIObject -ComputerName $Site_Server -Namespace root\sms\site_$Site_Code -class SMS_SUMDeploymentAssetDetails -Filter "AssignmentID = $DeploymentID" | select AssignmentID, AssignmentName
                    $DeploymentName = $DeploymentNameWithID.AssignmentName | select -Unique
                } else {
                    Write-Error "A Deployment ID was not specified. Aborting."
                    $HasErrors   = $True
                    throw   
            }
 
            if ($Status) {
                   $Output = Get-WMIObject -ComputerName $Site_Server -Namespace root\sms\site_$Site_Code -class SMS_SUMDeploymentAssetDetails -Filter "AssignmentID = $DeploymentID and StatusType = $StatusType" | `
                    select DeviceName, CollectionName, @{Name = 'StatusTime'; Expression = {$_.ConvertToDateTime($_.StatusTime) }}, @{Name = 'Status' ; Expression = {if ($_.StatusType -eq 1) {'Success'} elseif ($_.StatusType -eq 2) {'InProgress'} elseif ($_.StatusType -eq 5) {'Error'} elseif ($_.StatusType -eq 4) {'Unknown'}  }}
 
                } else {       
                    $Output = Get-WMIObject -ComputerName $Site_Server -Namespace root\sms\site_$Site_Code -class SMS_SUMDeploymentAssetDetails -Filter "AssignmentID = $DeploymentID" | `
                    select DeviceName, CollectionName, @{Name = 'StatusTime'; Expression = {$_.ConvertToDateTime($_.StatusTime) }}, @{Name = 'Status' ; Expression = {if ($_.StatusType -eq 1) {'Success'} elseif ($_.StatusType -eq 2) {'InProgress'} elseif ($_.StatusType -eq 5) {'Error'} elseif ($_.StatusType -eq 4) {'Unknown'}  }}
            }
 
            if (-not $Output) {
                Write-Error "A Deployment with ID: $($DeploymentID) is not valid. Aborting"
                $HasErrors   = $True
                throw
                 
            }
 
        } catch {
             
         
        } finally {
            if (($HasErrors -eq $false) -and ($Output)) {
                Write-Output ""
                Write-Output "Deployment Name: $DeploymentName"
                Write-Output "Deployment ID:   $DeploymentID"
                Write-Output ""
                Write-Output $Output | Sort-Object Status
            }
        }
    }
 
    END {}
 
}

function Get-CMModule {
    [CmdletBinding()]
    param()
    
    Try {
        Write-Verbose "Trying to import SCCM Module"
        Import-Module (Join-Path $(Split-Path $ENV:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -Verbose:$false
        Write-Verbose "Nice...imported the SCCM Module"
        }
    Catch 
        {
        Throw "Failure to import SCCM Cmdlets."
        } 
}

Get-CMModule

function Get-CMSiteCode {
    $CMSiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
    return $CMSiteCode
}

$sitecode = get-cmsitecode

$SetSiteCode = $sitecode+":"
Set-Location $SetSiteCode

Function Get-PatchTuesday ($Month,$Year)  
 { 
    $FindNthDay=2 #Aka Second occurence 
    $WeekDay='Tuesday' 
    $todayM=($Month).ToString()
    $todayY=($Year).ToString()
    $StrtMonth=$todayM+'/1/'+$todayY 
    [datetime]$StrtMonth=$todayM+'/1/'+$todayY 
    while ($StrtMonth.DayofWeek -ine $WeekDay ) { $StrtMonth=$StrtMonth.AddDays(1) } 
    $PatchDay=$StrtMonth.AddDays(7*($FindNthDay-1)) 
    return $PatchDay
    Write-Log -Message "Patch Tuesday this month is $PatchDay" -Severity 1 -Component "Set Patch Tuesday"
   Write-Output "Patch Tuesday this month is $PatchDay"
 } 

function Get-CMClientDeviceCollectionMembership {
    [CmdletBinding()]
    param (
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$SiteServer = (Get-WmiObject -Namespace root\ccm -ClassName SMS_Authority).CurrentManagementPoint,
        [string]$SiteCode = (Get-WmiObject -Namespace root\ccm -ClassName SMS_Authority).Name.Split(':')[1],
        [switch]$Summary,
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )

begin {}
process {
    Write-Verbose -Message "Gathering collection membership of $ComputerName from Site Server $SiteServer using Site Code $SiteCode."
    $Collections = Get-WmiObject -ComputerName $SiteServer -Namespace root/SMS/site_$SiteCode -Credential $Credential -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$ComputerName' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID"
    if ($Summary) {
        $Collections | Select-Object -Property Name,CollectionID
    }
    else {
        $Collections    
    }
    
}
end {}
}

#endregion


$todayDefault = Get-Date
$todayCompare = (get-date).ToString("yyyy-MM-dd")
$patchdayDefault = Get-PatchTuesday -Month $todayDefault.Month -Year $todayDefault.Year 
$patchdayCompare = (Get-PatchTuesday -Month $todayDefault.Month -Year $todayDefault.Year).tostring("yyyy-MM-dd")
$ReportdayCompare = ($patchdayDefault.AddDays(6)).tostring("yyyy-MM-dd")






if ($todayCompare -eq $ReportdayCompare)

{

    $UpdateStatus = Get-SCCMSoftwareUpdateStatus -DeploymentID 16777362

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


$errorvalue = ($UpdateStatus | Where-Object {($_.status -eq 'error')}).count


$successvalue = ($UpdateStatus | Where-Object {($_.status -eq 'success')}).count

$colletionname = $UpdateStatus.collectionname | Select-Object -First 1


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




#endregion

#Region Mailsettings


#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable=$false

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer=$SMTP

#port ([int], required)
$Port=$MailPortnumber

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From=[MimeKit.MailboxAddress]$MailFrom

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList=[MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]$MailTo)


#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$CCList=[MimeKit.InternetAddressList]::new()
#$CCList.Add([MimeKit.InternetAddress]$EmailToCC)



#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList=[MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")


#subject ([string], required)
$Subject=[string]"Serverpatchning $MailCustomer $monthname $year"

#text body ([string], optional)
#$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody=[string]$Body

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList=[System.Collections.Generic.List[string]]::new()
$AttachmentList.Add("$HTMLFileSavePath")
#$AttachmentList.Add("$CSVFileSavePath")

# Mailparameters
$Parameters=@{
    "UseSecureConnectionIfAvailable"=$UseSecureConnectionIfAvailable    
    #"Credential"=$Credential
    "SMTPServer"=$SMTPServer
    "Port"=$Port
    "From"=$From
    "RecipientList"=$RecipientList
    #"CCList"=$CCList
    #"BCCList"=$BCCList
    "Subject"=$Subject
    #"TextBody"=$TextBody
    "HTMLBody"=$HTMLBody
    "AttachmentList"=$AttachmentList
}

#endregion

#Region Send Mail

Send-MailKitMessage @Parameters

set-location $PSScriptRoot

