# PatchManagement

A collection of scripts to automate the process of keeping control over patch managment in MEMCM.
## Set-ScheduleTaskPatchTuesday.ps1
The script´s requirement
- Powershell 7
- Send-MailKitMessage
- Configuration Manager powershell module ( Run it on siteserver or client with console installed)

## Maintenance Windows Support Tool
![image](https://user-images.githubusercontent.com/16079354/209634515-5acea4d5-d02f-4252-ac93-54a57d74cf90.png)
The most easy way to create maintenance windows in MECM on collections is to use my application.
PreReq:
- Adminrights in MECM
- Run on siteserver or on client with ConfigMgr console installed.

## Mail with reports
I´ve created some scripts to be used to generate reports and send by mail to recipients.

More info in the folder <https://github.com/DambergC/PatchManagement/tree/main/MailUpdateStatus>
