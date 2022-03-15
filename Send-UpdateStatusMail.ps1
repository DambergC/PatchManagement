#
#.Synopsis
#  Creates a html-mail with list on downloaded and deployed patches from Microsoft in a
#  updatepackage selected in the script.
#.DESCRIPTION
#   Query MECM for updates downloaded to a specific update-package. In the config-file you select how many days
#   back from todays date you want to check.
#   Skripts require Powershell 7 and send-mailkitmessage module https://www.powershellgallery.com/packages/Send-MailKitMessage
#.EXAMPLE
#   Send-UpdateStatusMail.ps1
######################################################################
# Configuration for the script
######################################################################
$LimitFrom = '-10'
$SiteCode = '<Sitecode:>'
$UpdateGroupName = '<UpdateGroupName>'
$Emailfrom = '<Your no-replay address>
$EmailTo = '<Recipients mailgroupaddress>'
$EmailCustomer = '<the name of your customer>'
$EmailSmtp = '<your smtpserver>'
$mailport = '25'

#Calculate the numbers of days from todays date
$limit = (get-date).AddDays($LimitFromXML)

# Get the powershell module for MEMCM
if (-not(Get-Module -name ConfigurationManager)) {
    Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')
}

# Get the powershell module for Send-MailKitMessage
if (-not(Get-Module -name Send-Mailkitmessage)) {
    Install-Module -Name Send-MailKitMessage
}

#Get-InstalledModule

# To run the script you must be on ps-drive for MEMCM
Push-Location

Set-Location $SiteCode

# Array to collect result
$Result = @()

$resultat = Get-CMDeployment | Where-Object softwarename -like '*server*'

foreach ($item in $resultat) 


{
$object = New-Object -TypeName PSObject
$object | Add-Member -MemberType NoteProperty -Name 'Collection' -Value $item.Collectionname
$object | Add-Member -MemberType NoteProperty -Name 'Last date' -Value $item.SummarizationTime
$object | Add-Member -MemberType NoteProperty -Name 'UpdateGroupName' -Value $item.ApplicationName
$object | Add-Member -MemberType NoteProperty -Name 'Errors' -Value $item.NumberErrors
$object | Add-Member -MemberType NoteProperty -Name 'In Progress' -Value $item.NumberInProgress
$object | Add-Member -MemberType NoteProperty -Name 'Success' -Value $item.NumberSuccess
$object | Add-Member -MemberType NoteProperty -Name 'Targeted' -Value $item.NumberTargeted
$object | Add-Member -MemberType NoteProperty -Name 'Unknown' -Value $item.NumberUnknown
$result += $object

}


$Title = "Total numbers of deployments in list are " + $result.count 


# CSS HTML
$header = @"
<style>

    th {

        font-family: Arial, Helvetica, sans-serif;
        color: White;
        font-size: 14px;
        border: 1px solid black;
        padding: 3px;
        background-color: Black;

    } 
    p {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 12px;

    } 
    tr {

        font-family: Arial, Helvetica, sans-serif;
        color: black;
        font-size: 12px;
        vertical-align: text-top;

    } 

    body {
        background-color: white;
      }
      table {
        border: 1px solid black;
        border-collapse: collapse;
      }

      td {
        border: 1px solid black;
        padding: 5px;
        background-color: #e6f7d0;
      }

</style>
"@
######################################################################
# Text added to mail before list of patches
######################################################################
$pre = @"
<img src='cid:logo.png' height="50">
<p>This report list status for updatedeployments from site $SiteCodeFromXML owned by $EmailCustomer.</p>
<p>$title</p>
"@

######################################################################
# Text added to mail last 
######################################################################
$post = "<p>Report generated on $((Get-Date).ToString()) from <b><i>$($Env:Computername)</i></b></p>"

######################################################################
# Mail with pre and post converted to Variable later used to send with send-mailkitmessage
######################################################################
$StatusFound = $result | Sort-Object dateposted -Descending | ConvertTo-Html -Title "Downloaded patches" -PreContent $pre -PostContent $post -Head $header

## Mailsettings
# using module Send-MailKitMessage

#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable=$false

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer=$EmailSmtp

#port ([int], required)
$Port=$mailport

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From=[MimeKit.MailboxAddress]$Emailfrom

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList=[MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]$EmailTo)

#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$CCList=[MimeKit.InternetAddressList]::new()
$CCList.Add([MimeKit.InternetAddress]$EmailToCC)

#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList=[MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")


#subject ([string], required)
$Subject=[string]"$Emailcustomer - Status deployments $(get-date)"


#text body ([string], optional)
$TextBody=[string]"TextBody"

#HTML body ([string], optional)
$HTMLBody=[string]$StatusFound

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
######################################################################
#send message
######################################################################
Send-MailKitMessage @Parameters
