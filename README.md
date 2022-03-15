# PatchManagement
A collection of scripts to automate the process of keeping control over patch managment in MEMCM.
# Set-ScheduleTaskPatchTuesday.ps1
The script´s requirement
- Powershell 7
- Send-MailKitMessage
- Configuration Manager powershell module ( Run it on siteserver or client with console installed)
# Set-MaintenanceWindows.ps1
Script to Create one or more Maintance Windows for a Collection in MECM</b>

If you need to create one or more Maintance Windows in MECM for a Collection you can use this script.

You will have the following options- 
- CollID - CollectionID
- Offweek - How many weeks after patch tuesday
- Offdays - How many days after patch tuesday
- StartHour - The hour (0-23) when the Maintance Window should start
- StartMinutes - On the minute (0-59) when the Maintance Window should start
- EndHour - How long in hour for the Maintance Window
- EndMinutes - On the minute for the Maintance Window  
- Patchmount - Which month you want the Maintance Window (1-12)
- PatchYear - Which year
- ClearOldMW - If you want to clean up old Maintance Window for the Collection
- ApplyTo - If you want Any, TaskSequence or Only SoftwareUpdates to be controled by Maintance Window  

From your input the script will calculate Patch Tuesday for the month and set start- and endtime for the maintance Window.
# Send-UpdateStatusMail.ps1
  The script´s requirement
- Powershell 7.x
- Send-MailKitMessage
- Configuration Manager powershell module ( Run it on siteserver or client with console installed)
- The server or client where you configure to run the script need to be white-listed in your mailserver to be allowed to send mail.
# Send-UpdateDeployedMail.ps1
  The script´s requirement
- Powershell 7.x
- Send-MailKitMessage module
- Configuration Manager powershell module ( Run it on siteserver or client with console installed)

You need to edit row 246 to 276 in the script with your info to the recipients of the mail.
