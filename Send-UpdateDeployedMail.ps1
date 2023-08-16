#########################################################
#.Synopsis
#  List all updates member of updategroup in x-numbers of days
#.DESCRIPTION
#   Lists all assigned software updates in a configuration manager 2012 software update group that is selected 
#   from the list of available update groups or provided as a command line option
#.EXAMPLE
#   Send-UpdateDeployedMail.ps1
#
#   Skript created by Christian Damberg
#   christian@damberg.org
#   https://www.damberg.org
#
#########################################################
# Values need in script
#########################################################
#
# Numbers of days backwards you want to check for updates in updategroup
$LimitDays = '-15'
$SiteCode = '<sitecode>'
$UpdateGroupName = '<the name of your update group>'
$MailFrom = '<Your no-reply-address>'
$Mail_Error = '<Your mail to send when error happens>'
$Mail_Success = '<Your mail to the mailgroup>'
$MailSMTP = '<Your SMTP-Server>'
$MailPortnumber = '25'
$MailCustomer = 'Name of company'

#########################################################
# the function the extract the week number
#########################################################
function Get-ISO8601Week (){
    # Adapted from https://stackoverflow.com/a/43736741/444172
      [CmdletBinding()]
      param(
        [Parameter(
          ValueFromPipeline                =  $true,
          ValueFromPipelinebyPropertyName  =  $true
        )]                                           [datetime]  $DateTime
      )
      process {
        foreach ($_DateTime in $DateTime) {
          $_ResultObject   =  [pscustomobject]  @{
            Year           =  $null
            WeekNumber     =  $null
            WeekString     =  $null
            DateString     =  $_DateTime.ToString('yyyy-MM-dd   dddd')
          }
          $_DayOfWeek      =  $_DateTime.DayOfWeek.value__
    
          # In the underlying object, Sunday is always 0 (Monday = 1, ..., Saturday = 6) irrespective of the FirstDayOfWeek settings (Sunday/Monday)
          # Since ISO 8601 week date (https://en.wikipedia.org/wiki/ISO_week_date) is Monday-based, flipping Sunday to 7 and switching to one-based numbering.
          if ($_DayOfWeek  -eq  0) {
            $_DayOfWeek =    7
          }
    
          # Find the Thursday from this week:
          #     E.g.: If original date is a Sunday, January 1st     , will find     Thursday, December 29th     from the previous year.
          #     E.g.: If original date is a Monday, December 31st   , will find     Thursday, January 3rd       from the next year.
          $_DateTime                 =  $_DateTime.AddDays((4  -  $_DayOfWeek))
    
          # The above Thursday it's the Nth Thursday from it's own year, wich is also the ISO 8601 Week Number
          $_ResultObject.WeekNumber  =  [math]::Ceiling($_DateTime.DayOfYear    /   7)
          $_ResultObject.Year        =  $_DateTime.Year
    
          # The format requires the ISO week-numbering year and numbers are zero-left-padded (https://en.wikipedia.org/wiki/ISO_8601#General_principles)
          # It's also easier to debug this way :)
          $_ResultObject.WeekString  =  "$($_DateTime.Year)-W$("$($_ResultObject.WeekNumber)".PadLeft(2,  '0'))"
          Write-Output                  $_ResultObject
        }
      }
    }

#########################################################
# Section to extract monthname, year and weeknumbers
#########################################################

$monthname = (Get-Culture).DateTimeFormat.GetMonthName((get-date).month)
$year = (get-date).Year
$week = Get-ISO8601Week (get-date)
$nextweek = Get-ISO8601Week (get-date).AddDays(7)
$weeknumber = $week.weeknumber
$nextweeknumber = $nextweek.weeknumber

#########################################################
#Calculate the numbers of days from todays date
#########################################################
$limit = (get-date).AddDays($LimitDays)

#########################################################
# Get the powershell module for MEMCM 
# and Send-Mailkitmessage
#########################################################
if (-not(Get-Module -name ConfigurationManager)) {
    Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}

if (-not(Get-Module -name send-mailkitmessage)) {
    Install-Module send-mailkitmessage
    Import-Module send-mailkitmessage
}

#########################################################
# To run the script you must be on ps-drive for MEMCM
#########################################################
Push-Location
Set-Location $SiteCode

#########################################################
# Array to collect result
#########################################################
$Result = @()

#########################################################
# Gather all updates in updategrpup
#########################################################
$updates = Get-CMSoftwareUpdate -Fast -UpdateGroupName $UpdateGroupName

Write-host "Processing Software Update Group" $UpdateGroupName

forEach ($item in $updates)
{
    $object = New-Object -TypeName PSObject
    $object | Add-Member -MemberType NoteProperty -Name ArticleID -Value $item.ArticleID
    #$object | Add-Member -MemberType NoteProperty -Name BulletinID -Value $item.BulletinID
    $object | Add-Member -MemberType NoteProperty -Name Title -Value $item.LocalizedDisplayName
    $object | Add-Member -MemberType NoteProperty -Name LocalizedDescription -Value $item.LocalizedDescription
    $object | Add-Member -MemberType NoteProperty -Name DatePosted -Value $item.Dateposted
    $object | Add-Member -MemberType NoteProperty -Name Deployed -Value $item.IsDeployed
    $object | Add-Member -MemberType NoteProperty -Name 'URL' -Value $item.LocalizedInformativeURL
    $object | Add-Member -MemberType NoteProperty -Name 'Required' -Value $item.NumMissing
    $object | Add-Member -MemberType NoteProperty -Name 'Installed' -Value $item.NumPresent
    $object | Add-Member -MemberType NoteProperty -Name 'Severity' -Value $item.SeverityName
    $result += $object
}

#########################################################
# Create a row in the email to present numbers of updates
#########################################################
#$Numbersofupdates = "Totalt antal patchar frÃ¥n Microsoft som finns i Uppdateringspaketet " + $UpdateGroupName + " = " + $result.count
$Numbersofupdates = "Total numbers of updates from Microsoft that exist in updatepackage " + $UpdateGroupName + " = " + $result.count

#########################################################
# Create the list of updates sorted and only in the limit
#########################################################
$UpdatesFound = $result | Sort-Object dateposted -Descending | where-object { $_.dateposted -ge $limit }

#########################################################
# CSS HTML
#########################################################
$header = @"
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

################################################################
# Check if the downloaded updates in the limit are zero or not
################################################################
if ($UpdatesFound -eq $null )
{
    write-host "No updates downloaded or deployed since $limit"


    #Emailsettings when updates equals none
    $EmailTo = $Mail_Error
    
    

    $UpdatesFound = @"
    <br>
<img src='cid:logo.png' height="50">
<br>
    <B>No updates downloaded or deployed since $limit</B><br><br>
    <p></p>
<p>Action needed from third-line support</p>
"@
}

else 

{

#########################################################    
# Text added to mail before list of patches
#########################################################

#########################################################
# Emailsettings when updates more then one downloaded
#########################################################
$EmailTo = $Mail_Success

#########################################################
# The top of the email
#########################################################
$pre = @"
<br>
<img src='cid:logo.png' height="50">
<br>
<p><b>New updates!</b><br> 
<p>Updates will be available from wednesday week $weeknumber kl.15.00</p>
<p><b>Schema</b><br>
<p>The updates will be installed as follows:</p>
<p><ol>Test - Week $weeknumber - Every night between 03.00 - 08:00 (If any updates are published)</ol></p>
<p><ol>Prod - Week $nextweeknumber - Majority will be installed saturday 11.00pm till Sunday 09.00am</ol></p>
<p><ol>AX - Managed manually by the administration.</ol></p>
<p><b>Patchar From Microsoft</b><br>
<p>The following updates are downloaded and published in updategroup <b><i>$UpdateGroupName</i></b> since $limit</p>
<p>$Numbersofupdates</p>
"@

#########################################################
# Footer of the email
#########################################################
$post = @"
<p>Raport created $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>
<p>Script created by:<br><a href="mailto:Your Email">Your name</a><br>
<a href="https://your blog">your description of your blog</a>
"@

##########################################################################################
# Mail with pre and post converted to Variable later used to send with send-mailkitmessage
##########################################################################################
$UpdatesFound = $result | Sort-Object dateposted -Descending | where-object { $_.dateposted -ge $limit }| ConvertTo-Html -Title "Downloaded patches" -PreContent $pre -PostContent $post -Head $header

}

#########################################################
# Mailsettings
# using module Send-MailKitMessage
#########################################################

#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable=$false

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer=$MailSMTP

#port ([int], required)
$Port=$MailPortnumber

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From=[MimeKit.MailboxAddress]$MailFrom

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList=[MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]$MasailTo)


#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$CCList=[MimeKit.InternetAddressList]::new()
#$CCList.Add([MimeKit.InternetAddress]$EmailToCC)



#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList=[MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")

# Different subject depending on result of search for patches.
if ($UpdatesFound -ne $null )
{
#subject ([string], required)
$Subject=[string]"Serverpatchning $MailCustomer $monthname $year"
}
else 
{
#subject ([string], required)
$Subject=[string]"Error Error - Action needed $(get-date)"    
}

#text body ([string], optional)
#$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody=[string]$UpdatesFound

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList=[System.Collections.Generic.List[string]]::new()
$AttachmentList.Add("$PSScriptRoot\logo.png")

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
#########################################################
#send email
#########################################################
Send-MailKitMessage @Parameters
