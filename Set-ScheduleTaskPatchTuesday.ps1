<#
.SYNOPSIS
This script creates x-numbers of scheduled task based on when patch tuesday occurs.
    
.DESCRIPTION
This script give you the option to create multiple scheduled task to send a mail x-days after patch tuesday.

############################################################
Christian Damberg
www.damberg.org
Version 1.0
2021-12-22
############################################################
    
.EXAMPLE
.\Set-ScheduleTaskPatchTuesday.ps1 -OffSetWeeks 0 -OffSetDays 2 -AddStartHour 12 -AddStartMinutes 0 -PatchMonth "1","2","3","4","5","6","7","8","9","10","11" -patchyear 2022 -FolderName PatchMail -execute 'pwsh.exe' -scriptpath 'C:\Scripts\PatchManagement\Send-UpdateDeployedMail.ps1' -UserName 'Korsberga.local\scriptrunner' -Verbose
Will create Schedule Task for January to November to send mail two days after patch tuesday at noon.

.DISCLAIMER
All scripts and other Powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>
############################################################
# Parameters
############################################################

PARAM(
    [int]$OffSetWeeks,
    [int]$OffSetDays,
    [Parameter(Mandatory=$True)]
    [int]$AddStartHour,
    [Parameter(Mandatory=$True)]
    [int]$AddStartMinutes,
    [string[]]$PatchMonth,
    [Parameter(Mandatory=$True)]
    [int]$patchyear,
    [string]$FolderName,
    [string]$UserName,
    [string]$execute,
    [string]$scriptpath
    )  

############################################################
# region functions
############################################################


$password = read-host -Prompt "The domain password for $username"


# Read passwordfile
#$Encrypted = Get-Content $PathPasswordFile | ConvertTo-SecureString
#$UnsecurePassword = (New-Object PSCredential "user",$encrypted).GetNetworkCredential().Password

# Create variable with username and password
#$Credential = New-Object System.Management.Automation.PsCredential($UserName, $Encrypted)


# Set Patch Tuesday for a Month 
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
Set-Location $PSScriptRoot

# Function for append events to logfile located c:\windows\logs
Function Write-Log
{
    PARAM(
    [String]$Message,
    [int]$Severity,
    [string]$Component
    )
    Set-Location $PSScriptRoot
    $Logpath = "C:\Windows\Logs"
    $TimeZoneBias = Get-CimInstance win32_timezone
    $Date= Get-Date -Format "HH:mm:ss.fff"
    $Date2= Get-Date -Format "MM-dd-yyyy"
        "<![LOG[$Message]LOG]!><time=$([char]34)$Date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$Component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$Severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath "$Logpath\Create-ScheduleTask.log" -Append -NoClobber -Encoding default

}

# Function to create a folder in Scheduled Task
Function New-ScheduledTaskFolder
    {
     Param ($taskpath)

     $ErrorActionPreference = "stop"
     $scheduleObject = New-Object -ComObject schedule.service
     $scheduleObject.connect()
     $rootFolder = $scheduleObject.GetFolder("\")
        Try {$null = $scheduleObject.GetFolder($taskpath)}
        Catch { $null = $rootFolder.CreateFolder($taskpath) }
        Finally { $ErrorActionPreference = "continue" } }

# Misc variables needed in script
$ErrorMessage = $_.Exception.Message
$MonthArray = New-Object System.Globalization.DateTimeFormatInfo 
$MonthNames = $MonthArray.MonthNames 

# Create a subfolder in Scheduled Task
New-ScheduledTaskFolder $FolderName

# Create Scheduled Task for for every month specified in variable Patchmonth
foreach ($Monthnumber in $PatchMonth) 
{
    # Set Patch Tuesday for each Month 
    $PatchDay = Get-PatchTuesday $Monthnumber $PatchYear
                 
    # Set month number correct to display name later in script (Months array starting from 0 hence the -1) 
    $displaymonth = $Monthnumber - 1 

    # Set starttime for schedule task
    $StartTime=$PatchDay.AddDays($OffSetDays).AddHours($AddStartHour).AddMinutes($AddStartMinutes)
    
    ############################################################
    # This section must be edited before running the script
    ############################################################
    # Action in Scheduled Task
    $taskAction = New-ScheduledTaskAction `
    -Execute $execute `
    -Argument "-File $scriptpath -ExecutionPolicy bypass"
    ############################################################
    # Done
    ############################################################
    # Create a new trigger (Daily at 3 AM)
    $tasktrigger = New-ScheduledTaskTrigger -At $StartTime -Once

    # The name of your scheduled task.
    $taskName = "Patchstatus-Mail " +$MonthNames[$displaymonth] + " "+ $patchyear

    # Describe the scheduled task.
    $description = "Mail - Status on downloaded and deployed patches"

    $Taskusername = $Credential.UserName
    $TaskPwd = $Credential.Password

        Try 
        {
            # Register the scheduled task
            Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger $taskTrigger -Description $description -TaskPath $FolderName -User $username -Password $password -RunLevel Highest

            #Write-Log -Message "Created schedule task $taskname " -Severity 1 -Component "New Schedule Task"
          
        }
        Catch 
        {
            Write-Warning "$_.Exception.Message"
            Write-Log -Message "$_.Exception.Message" -Severity 3 -Component "Create Schedule Task"
        }

    }

Set-Location $PSScriptRoot
