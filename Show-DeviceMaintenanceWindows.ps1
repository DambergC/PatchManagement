
# Functions
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

# Array
$ResultColl = @()

# Devices
$devices = Get-CMCollectionMember -CollectionId SMSDM003

$today = Get-Date
$checkdatestart = $today.AddDays(-10)
$checkdateend = $today.AddDays(30)
$filedate = get-date -Format yyyMMdd

# Loop for each device
foreach ($device in $devices)
        
        {

            # Get all Collections for Device
            $collectionids = Get-CMClientDeviceCollectionMembership -ComputerName $device.name

            # Check every Collection for service windows
            foreach ($collectionid in $collectionids)
            {
                # Only include Collections with Service Windows
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
                                #$object | Add-Member -MemberType NoteProperty -Name 'CollectionID' -Value $collectionid.collectionid
                                $object | Add-Member -MemberType NoteProperty -Name 'Collection-Name' -Value $collectionid.name

                                $object | Add-Member -MemberType NoteProperty -Name 'MW-Name' -Value $mw.name
                                $object | Add-Member -MemberType NoteProperty -Name 'StartTime' -Value $mw.StartTime
                                $object | Add-Member -MemberType NoteProperty -Name 'Duration' -Value $mw.Duration
          
                                $resultColl += $object
                                }
                            }
        
                    }
            }

        }





New-HTML -TitleText 'Maintenance Windows - Kriminalvården' -FilePath "c:\temp\KVV_MW_$filedate.HTML" -ShowHTML -Online {

    New-HTMLHeader {
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLText -Text 'Kriminalvården - Maintenance Windows' -FontSize 35 -Color Darkblue -FontFamily Arial -Alignment center
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
