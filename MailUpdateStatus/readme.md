# MailUpdateStatus

I created these scripts to facilitate the work of monitoring the status of Windows Updates via ConfigMgr.

With the script, you don't have to sit and work more than necessary when it comes to keeping track of the patches every month.
## scriptconfig.xml
The first file in the folder is "scriptconfig.xml" which contains everything the scripts need to run. With that file, you don't have to open and edit any of the script more than necessary, or rather only if you can rename "scriptconfig.xml" to something else or put it somewhere else...but why ?
|XML Element|explanation|
## Send-WindowsUpdateDeployed.ps1
This script runs a check against an Update Group in ConfigMgr and retrieves all patches that have been published in the last x-numbers of days (the value is in scriptconfig.xml)
## Send-WindowsUpdateStatus.ps1
The script produces a report for a Windows Update Deployment in ConfigMgr which then compiles a report on which devices are Compliant and the report is sent via e-mail.

The script can work with multiple deployments. Values ​​for which deployment and x-number of days after Patch Tuesday are controlled via "scriptconfig.xml"
## Extra in script
Extra i skripten är att det finns funktion för loghantering och det som jag lagt till är att när logfilen som tillhör skripten kommer över x-antal bytes (2000000 default) så byter den  namn och flyttas. Detta för att underlätta hanteringen av logfilen.
