

# Array
$ResultColl = @()

# Devices
$devices = Get-CMCollectionMember -CollectionId SMSDM003

$today = Get-Date
$checkdatestart = $today.AddDays(10)
$checkdateend = $today.AddDays(30)

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

            # Show result
            $resultColl  | Sort-Object device,collectionid,starttime | Format-Table

            $ResultColl | Out-HtmlView -FilePath c:\temp\ttt.html -PreventShowHTML -Title "Kriminalvården patch status $today" -PagingLength 100

