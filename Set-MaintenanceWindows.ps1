<#
.SYNOPSIS
This script configures multiple Maintenance Windows for a collection.The schedule is based on offset-settings with Patch Tuesday as base.
    
.DESCRIPTION
This script give you options to delete existing Maintenance Windows on collection, decide if the Maintenance Windows should be for Any installation, Task sequence
or only SoftwareUpdates.

You can also decide which month the Maintenance Windows should be configured for.

Some of the funcionality has been borrowed from Daniel Engberg´s script, created 2018 which he borrowed som functionality from Octavian Cordos' script, created in 2015.

####################
Christian Damberg
www.damberg.org
Version 1.0
2021-12-22
####################
    
.EXAMPLE
.\Set-Maintenance.ps1 -CollID ps100137 -OffSetWeeks 1 -OffSetDays 5 -AddStartHour 18 -AddStartMinutes 0 -AddEndHour 4 -AddEndMinutes 0 -PatchMonth "1","2","3","4","5","6","7","8","9","10","11" -patchyear 2022 -ClearOldMW Yes -ApplyTo SoftWareUpdatesOnly
Will create a Maintenance Window with Patch Tuesday + 1 week and 5 days for collection with ID PS100137 for every month except december in 2022. The script also delete old Maintance Windows and the new Maintance Windows are only for SoftwareUpdates.
    
.DISCLAIMER
All scripts and other Powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>


PARAM(
    [int]$OffSetWeeks,
    [int]$OffSetDays,
    [Parameter(Mandatory=$True)]
    [int]$AddStartHour,
    [Parameter(Mandatory=$True)]
    [int]$AddStartMinutes,
    [Parameter(Mandatory=$True)]
    [int]$AddEndHour,
    [Parameter(Mandatory=$True)]
    [int]$AddEndMinutes,
    [Parameter(Mandatory=$True)]
    [string[]]$PatchMonth,
    [Parameter(Mandatory=$True)]
    [int]$patchyear,
    [Parameter(Mandatory=$True)]
    [ValidateSet('Yes','No')]
    [string]$ClearOldMW = "No", 
    [Parameter(Mandatory=$True)]
    [ValidateSet('SoftWareUpdatesOnly','TaskSequenceOnly','Any')]
    [string]$ApplyTo = "SoftWareUpdatesOnly",
    [Parameter(Mandatory=$True,Helpmessage="Provide the ID of the collection")]
    [string]$CollID
    )  

#region Initialize

#Load SCCM Powershell module
Try {
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -ErrorAction Stop
}
Catch [System.UnauthorizedAccessException] {
    Write-Warning -Message "Access denied" ; break
}
Catch [System.Exception] {
    Write-Warning "Unable to load the Configuration Manager Powershell module from $env:SMS_ADMIN_UI_PATH" ; break
}

#endregion

#region functions

#Set Patch Tuesday for a Month 
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

 }  
 
#Remove all existing Maintenance Windows for a Collection 
Function Remove-MaintenanceWindow 
{
    PARAM(
    [string]$CollID
    )
    $SiteCode = Get-PSDrive -PSProvider CMSITE
    Set-Location -Path "$($SiteCode.Name):\"

    $OldMW = Get-CMMaintenanceWindow -CollectionId $CollID

    #Get-CMMaintenanceWindow -CollectionId $CollID | Where-Object {$_.StartTime -lt (Get-Date)} | ForEach-Object { 
    Get-CMMaintenanceWindow -CollectionId $CollID | ForEach-Object { 
    Try {
        Remove-CMMaintenanceWindow -CollectionID $CollID -Name $_.Name -Force -ErrorAction Stop
        $Coll=Get-CMDeviceCollection -CollectionId $CollID -ErrorAction Stop
        Write-Log -Message "Removing $($_.Name) from collection $MWCollection" -Severity 1 -Component "Remove Maintenance Window"
        Write-Output "Removing $($_.Name) from collection $MWCollection"
    }
    Catch {
        Write-Log -Message "Unable to remove $($_.Name) from collection $MWCollection" -Severity 3 -Component "Remove Maintenance Window"
        Write-Warning "Unable to remove $($_.Name) from collection $MWCollection. Error: $_.Exception.Message"   
    } 
}
Set-Location $PSScriptRoot
} 

#Function for append events to logfile located c:\windows\logs
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
    $Type=1
    "<![LOG[$Message]LOG]!><time=$([char]34)$Date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$Component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$Severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath "$Logpath\Set-MaintenanceWindows.log" -Append -NoClobber -Encoding default
    $SiteCode = Get-PSDrive -PSProvider CMSITE
    Set-Location -Path "$($SiteCode.Name):\"
}

#endregion

#region Parameters

$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location -Path "$($SiteCode.Name):\"
$GetCollection = Get-CMDeviceCollection -ID $CollID
$MWCollection = $GetCollection.Name
$ErrorMessage = $_.Exception.Message

#endregion


#Delete old Maintance Windows if value in $ClearOldMW equals Yes
if ($ClearOldMW -eq 'Yes')
{

    Remove-MaintenanceWindow -CollID $CollID
    
}



#Create Maintance Windows for for every month specified in variable Patchmonth
foreach ($Monthnumber in $PatchMonth) 

{

$MonthArray = New-Object System.Globalization.DateTimeFormatInfo 
$MonthNames = $MonthArray.MonthNames 

#Set Patch Tuesday for each Month 
$PatchDay = Get-PatchTuesday $Monthnumber $PatchYear

#Fix to get the right Name of Maintance Windows month.
$displaymonth = $Monthnumber - 1 
                 
#Set Maintenance Window Naming Convention (Months array starting from 0 hence the -1) 
$NewMWName =  "MW_"+$MonthNames[$displaymonth]+"_"+$patchyear
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location -Path "$($SiteCode.Name):\"

#Set Device Collection Maintenace interval  
$StartTime=$PatchDay.AddDays($OffSetDays).AddHours($AddStartHour).AddMinutes($AddStartMinutes)
$EndTime=$StartTime.Addhours(0).AddHours($AddEndHour).AddMinutes($AddEndMinutes)

Try {
    #Create The Schedule Token  
    $Schedule = New-CMSchedule -Nonrecurring -Start $StartTime.AddDays($OffSetWeeks*7) -End $EndTime.AddDays($OffSetWeeks*7) -ErrorAction Stop
    
    New-CMMaintenanceWindow -CollectionID $CollID -Schedule $Schedule -Name $NewMWName -ApplyTo $ApplyTo -ErrorAction Stop
    Write-Log -Message "Created Maintenance Window $NewMWName for Collection $MWCollection" -Severity 1 -Component "New Maintenance Window"
    #Write-Output "Created Maintenance Window $NewMWName for Collection $MWCollection" 
}
Catch {
    Write-Warning "$_.Exception.Message"
    Write-Log -Message "$_.Exception.Message" -Severity 3 -Component "Create new Maintenance Window"
}

}

Set-Location $PSScriptRoot
