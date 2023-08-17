<#
-------------------------------------------------------------------------------------------------------------------------
.Synopsis
   Generate htmlpage with Devices and Maintenance Windows
.DESCRIPTION
   Script to be run as schedule task on siteserver. It's recommended to be use my script to
   Generate scheduleTask based on offset from patchTuesday.

   https://github.com/DambergC/PatchManagement/blob/main/Set-ScheduleTaskPatchTuesday.ps1

   The script generate a html-page and if you use the send-mailkitmessage it will send a mail
   to a group of administrators with info about the Maintenace Windows for a devices in a 
   collection.
.EXAMPLE
   Show-DeviceMaintenanceWindows.ps1
.VERSION
   0.0.9    Development only
.DISCLAIMER
All scripts and other Powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
-------------------------------------------------------------------------------------------------------------------------
#>

#Region Parameters

# Date section
$today = Get-Date
$checkdatestart = $today.AddDays(-10)
$checkdateend = $today.AddDays(30)
$filedate = get-date -Format yyyMMdd
$TitleDate = get-date -DisplayHint Date
$counter = 0
$HTMLFileSavePath = "c:\temp\KVV_MW_$filedate.HTML"
$HTMLHeadline = "Kriminalvården - Maintenance Windows $TitleDate"
$SMTP = 'smtp.kvv.se'
$MailFrom = 'no-reply@kvv.se'
$MailTo = 'christian.damberg@kriminalvarden.se'
$MailPortnumber = '25'
$MailCustomer = 'Kriminalvården - IT'
#$collectionidToCheck = 'PS100056'
$collectionidToCheck = 'PS10007B' <# TEST TEST TEST #>
$siteserver = 'vntsql0081'

#endregion

#region modules

<# 
-------------------------------------------------------------------------------------------------------------------------
Required Modules (installed offline under c:\program files\Windowspowershell\Modules

Guide - https://johnnycase.github.io/post/2021/05/17/pwrshl-module-offline.html

Send-MailkitMessage - https://www.powershellgallery.com/packages/Send-MailKitMessage/3.2.0

psWriteHTML - https://www.powershellgallery.com/packages/PSWriteHTML/1.2.0
-------------------------------------------------------------------------------------------------------------------------
#>
Import-Module send-mailkitmessage
import-module PSWriteHTML

#endregion

#Region Functions needed in script

# Get cmmodule and install it
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

# Get the sitecode
function Get-CMSiteCode {
    $CMSiteCode = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer | Select-Object -ExpandProperty SiteCode
    return $CMSiteCode
}

# Get the month patchtuesday
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

# Get all collections for a Device
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

#Region Script part 1 collect info from selected collection and check devices membership in Collections with Maintenance Windows

# Array to collect data in
$ResultColl = @()

# Devices
$devices = Get-CMCollectionMember -CollectionId $collectionidToCheck


# For the progressbar
$complete = 0

$scriptstart = (get-date).Second

# Loop for each device
foreach ($device in $devices)
        
        {
            $counter++
            Write-Progress -Activity 'Processing computer' -CurrentOperation $device.Name -PercentComplete (($counter / $devices.count)*100)
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
                                # Only show Maintenance Windows waiting to run
                                if ($mw.StartTime -gt $checkdatestart -and $mw.StartTime -lt $checkdateend)
                                {

                                $object = New-Object -TypeName PSObject
                                $object | Add-Member -MemberType NoteProperty -Name 'Device' -Value $device.name
                                $object | Add-Member -MemberType NoteProperty -Name 'Collection-Name' -Value $collectionid.name
                                $object | Add-Member -MemberType NoteProperty -Name 'StartTime' -Value $mw.StartTime
                                $object | Add-Member -MemberType NoteProperty -Name 'Duration' -Value $mw.Duration
          
                                $resultColl += $object
                                }
                            }
                    }
            }
        }

$scriptstop = (get-date).Second

#endregion

#region Script part 2 Create the html-file to be distributed

New-HTML -TitleText "Maintenance Windows - Kriminalvården" -FilePath $HTMLFileSavePath -ShowHTML -Online {

    New-HTMLHeader {
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLText -Text $HTMLFileHeadLine -FontSize 35 -Color Darkblue -FontFamily Arial -Alignment center
                New-HTMLHorizontalLine
            }
        }
    }

    New-HTMLSection -Invisible -Title "Maintenance Windows $filedate"{

        New-HTMLTable -DataTable $ResultColl -PagingLength 25
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

#endregion

#Region CSS and HTML for mail thru Send-MailKitMessage

$style = @"
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
    tr {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 11px;
        vertical-align: text-top;

    } 

    body {
        background-color: #B1E2EC;
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
"@

#endregion

#Region HTML Mail

$header = @"

<p><b>Server Maintenance Windows - List</b><br> 

"@


$Body = @"

<p><b>Script runtime ($scriptstop - $scriptstart) seconds</b><br></p> 

"@


$post = @"
<p>Report created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
<p>Script created by:<br><a href="mailto:Your Email">Your name</a><br>
<a href="https://your blog">your description of your blog</a>
"@

ConvertTo-Html -Title "rrrrrrr" -PreContent $pre -PostContent $post -Head $header -Body $body

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

#endregion
