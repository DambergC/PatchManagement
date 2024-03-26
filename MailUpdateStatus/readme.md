# MailUpdateStatus

I created these scripts to facilitate the work of monitoring the status of Windows Updates via ConfigMgr.

With the script, you don't have to sit and work more than necessary when it comes to keeping track of the patches every month.
## scriptconfig.xml
The first file in the folder is "scriptconfig.xml" which contains everything the scripts need to run. With that file, you don't have to open and edit any of the script more than necessary, or rather only if you can rename "scriptconfig.xml" to something else or put it somewhere else...but why ?
| XML-element | Explanation |
| ------ | ------ |
| Logfile\path | Path to logfile |
| Logfile\Name | logfilename |
| Logfile\Lofilethrehold | Max size for logfile before rotation |
| HTMLFilePath | Where the script create and store html-files to be attached in mail |
| RunScript | Your Deployments (DeploymentID, Offsetdays, Description) |
| DisableReportMonth | If you don´t want the script to run on a specific month...why? |
| Recipients | Who do you want to send the report to? |
| UpdateDeployed\LimitDays | Number of days back in time to check for published updates from Microsoft |
| UpdateDeployed\UpdateGroupName | The name on the UpdateGroup where you deployed the updates |
| UpdateDeployed\DaysAfterPatchToRun | Number of days after Patch Tuesday to run the script |

## Send-WindowsUpdateDeployed.ps1
This script runs a check against an Update Group in ConfigMgr and retrieves all patches that have been published in the last x-numbers of days (the value is in scriptconfig.xml)
## Send-WindowsUpdateStatus.ps1
The script produces a report for a Windows Update Deployment in ConfigMgr which then compiles a report on which devices are Compliant and the report is sent via e-mail.

The script can work with multiple deployments. Values ​​for which deployment and x-number of days after Patch Tuesday are controlled via "scriptconfig.xml"
## Extra in script
In the scripts, there is a function for log management and what I added is that when the log file that is created and written in by the script exceeds x-number of bytes (2000000 bytes default), it is renamed and moved. This is to facilitate the handling of the log file.
