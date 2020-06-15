## Helper functions
function Get-VCConnectionState {
    # Are we currently connected to anything?
    If ($global:DefaultVIServer) {
        # Ask if want to stay connected
        Write-Host "Already connected to $global:DefaultVIServer" -ForegroundColor Green
        $StayConnected = Read-Host "Stay connected - Y/N"
        
        # If reply is to stay connected, return from the function
        If ($StayConnected -eq "Y") {
            Return 
        } else {
            # Disconnect from this session
            Disconnect-VIServer -Confirm:$false
        }
    }

    # We're not connected
    $targetVC = Read-Host "Enter the VC to connect to"
    # In the event that we fail to connect, try a maximum of 3 times 
    $Attempts = 0
    If (-not $global:DefaultVIServer) {
        do {
            try {
                Connect-VIServer $targetVC -ErrorAction stop
                Return
            } catch {
                $Attempts++
                Write-Host "Failed to connect $Attempts time(s)." -ForegroundColor Green
            }
        } while ($Attempts -le 2)
        Write-Host "Too many failures" -ForegroundColor Green
        Read-Host "Press ENTER to exit the script"
        Break
    } 
} # end function Get-VCConnectionState

Function Set-InitializeScript {
    $temp_home = [Environment]::GetFolderPath("MyDocuments") # Get the MyDocuments folder - we'll be creating a folder here for output. Using environment supports roaming profile path.
    $Global:SnapshotManagement_home = $temp_home + "\SnapshotManagementReports" # This is going to be our default for logging output. Remainder will be subfolders to this.
    
    # If the folder for logging output doesn't exist, create it and the subfolders.
    Clear-Host
    Write-Host "Initialization checks in progress." -ForegroundColor Green
    Write-Host "----------------------------------" -ForegroundColor Green
    Write-Host "`nChecking if input and report folders exist." -ForegroundColor Green
    If (!(Test-Path -Path $Global:SnapshotManagement_home)) {
        Write-Host "Folders not present. Creating." -ForegroundColor Green
        New-Item -ItemType directory -Path $Global:SnapshotManagement_home
    } else {
        Write-Host "Folder present." -ForegroundColor Green
    }
    
    Write-Host "`nCheck if ImportExcel module is available." -ForegroundColor Green
    Test-ForImportExcel
        
    Write-Host "`nInitialization checks complete." -ForegroundColor Green
    Read-Host "`nPress ENTER to continue."
} # End function Set-InitializeScript

# Report generating use the ImportExcel module available from https://www.powershellgallery.com/packages/ImportExcel/7.1.0
# It only tests if it's avaialble or not - should probably be extended to actually offer to install if not available.
Function Test-ForImportExcel {
    $CurrentlyAvailableModules = (Get-Module -ListAvailable | Where-Object { $_.Name -eq "ImportExcel" })
    If ($CurrentlyAvailableModules) {
        Write-Host "ImportExcel module is available" -ForegroundColor Green
        Return $True
    } else {
        Return $False
    }
} # End function Check-ForImportExcel

function Check-VMExists ($targetVM) {
    Try {
        Get-VM $targetVM  -ErrorAction Stop | Out-Null
    } Catch {
        Write-Host "`nVM $targetVM doesn't exist." -ForegroundColor Red
        Read-Host "Press ENTER to return to main menu."
        #Clear-Host
        SnapshotManagementMenu
    }
} # end Function Check-VMExists

# Slightly modified version of https://enterpriseadmins.org/blog/scripting/get-vcenter-scheduled-tasks-with-powercli-part-1/
Function Get-VIScheduledTasks {
    (Get-View ScheduledTaskManager).ScheduledTask | 
        ForEach-Object { (Get-View $_ -Property Info).Info } | 
        Select-Object Name, Description, Enabled, Notification, LastModifiedUser, State, Entity,
    @{N = "Entity Name"; E = { (Get-View $_.Entity -Property Name).Name } },
    @{N = "Last Modified Time"; E = { $_.LastModifiedTime } },
    @{N = "Next Run Time"; E = { $_.NextRunTime } },
    @{N = "Prev Run Time"; E = { $_.LastModifiedTime } },
    @{N = "Action Name"; E = { $_.Action.Name } }
} # end Function Get-VIScheduledTasks

function Send-EmailReport ($Subject, $ReportToSend) {
    # Fill in as necessary
    $FromUser = "someone@someplace"
    $ToUser = "someone@someplace"
    $mailServer = "xxx.xxx.xxx.xxx"

    Send-MailMessage -From $FromUser -To $ToUser -Subject $Subject -SmtpServer $mailServer -Attachments $ReportToSend
} # end function Send-EmailReport

## Functions for the actual menu choices.

function Get-ListOfVMsWithSnapshots {
    #Clear-Host
    Write-Host "List all VMs with snapshots." -ForegroundColor Green
    Write-Host "----------------------------" -ForegroundColor Green
    Write-Host "This returns a simple list of all the VMs with a snapshot." -ForegroundColor Green
    Write-Host "If a VM appears more than once, it has multiple snapshots." -ForegroundColor Green
    Write-Host "For a more detiled report, chose the detailed report option." -ForegroundColor Green

    Get-VM | Get-Snapshot | Select-Object VM | Format-Table -AutoSize
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end function Get-ListOfVMsWithSnapshots

function Get-SnapshotReport {
    #Clear-Host
    Write-Host "Generate detailed report of all snapshots." -ForegroundColor Green
    Write-Host "------------------------------------------" -ForegroundColor Green
    $snapshotCollection = @()

    $dcs = Get-Datacenter 
    foreach ($dc in $dcs) {

        Write-Host "`nProcessing Snapshots information in datacenter : $dc." -ForegroundColor Green
        foreach ($snap in Get-VM -Location $dc | Get-Snapshot) {
            $ds = Get-Datastore -VM $snap.vm
            $datastorePercentageFree = $ds | Select-Object @{N = "PercentFree"; E = { [math]::round($_.FreeSpaceGB / $_.CapacityGB * 100) } }
            $SnapshotAge = ((Get-Date) - $snap.Created).Days
        
            $snapinfo = [PSCustomObject]@{
                "vCenter"                   = $vcName
                "VM"                        = $snap.vm
                "Snapshot Name"             = $snap.name
                "Current snapshot?"         = $snap.IsCurrent
                "Description"               = $snap.description
                "Created"                   = $snap.created
                "Snapshot age (days)"       = $SnapshotAge
                "Snapshot size (GB)"        = [math]::round($snap.sizeGB)
                "Datastore"                 = $ds[0].name
                "Datastore free space (GB)" = [math]::round($ds[0].FreeSpaceGB)
                "Datastore percent free (%)" = $datastorePercentageFree.PercentFree
                "Current snapshot"          = $snap.IsCurrent
                "Memory state"              = $snap.Powerstate
                "Quiesced"                  = $snap.Quiesced
            } # end $snapinfo = [PSCustomObject]@
            $snapshotCollection += $snapinfo
        }
        $snapshotCollection | Out-Host
    }

    If ($snapshotCollection.count -eq 0) {
        Write-Host "`nNo VMs with snapshots found." -ForegroundColor Green
    } else {
        $GenerateReport = Read-Host "Do you want to export this report to Excel - Y/N?"
        If ($GenerateReport -eq "Y") {
            $date = Get-Date -Format "yyyy-MMM-dd-HHmmss"
            $xlsx_output_file = "$Global:SnapshotManagement_home\$global:DefaultVIServer-SnapshotReport-$Date.xlsx"     
            $snapshotCollection | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname Snapshots -AutoSize 
            Write-Host "`nReport generated : $xlsx_output_file" -ForegroundColor Green

            $subj = "Snapshot audit report ..."
            Write-Host "Sending email report" -ForegroundColor Green
            Send-EmailReport $subj $xlsx_output_file        
        }
    }
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu

} # end function Get-SnapshotReport

Function Check-VMForSnapshot {
    $targetVM -eq $null
    #Clear-Host
    Write-Host "Check for snaphosts on a specified VM." -ForegroundColor Green
    Write-Host "--------------------------------------" -ForegroundColor Green

    $TargetVM = Read-Host "Enter the VM to check for existing snapshots."

    Check-VMExists $TargetVM
    $SingleVMSnapshotDetails = Get-VM $TargetVM | get-snapshot | Select-Object Name, Description, Created, SizeGB, IsCurrent, Quieseced, Powerstate, Children 
    If (($SingleVMSnapshotDetails).count -eq 0) {
        Write-Host "`nNo snapshot present." -ForegroundColor Green
    } else {
        Write-Host "VM has " ($SingleVMSnapshotDetails).count "snapshots."  
        Write-Host "`nSnapshot Details : `b"  -ForegroundColor Green 
        $SingleVMSnapshotDetails | Out-Host
    }
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end function Check-VMForSnapshot

function Get-VMsNeedingConsolidation {
    #Clear-Host
    Write-Host "Generate report of all VMs needing consolidation." -ForegroundColor Green
    Write-Host "-------------------------------------------------" -ForegroundColor Green

    Write-Host "`nThis will list the VMs that are marked as needing consolidation." -ForegroundColor Green
    Write-Host "If there are none requiring consolidation, the list will be empty." -ForegroundColor Green
    Write-Host "We are checking the value of ExtensionData.Runtime.ConsolidationNeeded for every VM." -ForegroundColor Green
    Write-Host "TRUE = consolidation required." -ForegroundColor Green
    Write-Host "FALSE = no consolidation required.`n" -ForegroundColor Green

    $snapshotConsolidationCollection = @()

    $dcs = Get-Datacenter
    ForEach ($dc in $dcs) {
        ForEach ($vm in Get-VM -Location $dc) {
            If ($vm.ExtensionData.Runtime.ConsolidationNeeded) {
                $snapinfo = [PSCustomObject]@{
                    "VM"                        = $vm.name
                    "Consolidation needed"      = $vm.ExtensionData.Runtime.ConsolidationNeeded
                }
                $snapshotConsolidationCollection += $snapinfo
            } 
        }
    }

    If ($snapshotConsolidationCollection.count -eq 0) {
        Write-Host "`nNo VMs with snapshots needing consolidation found." -ForegroundColor Green
    } else {
        $snapshotConsolidationCollection | Out-Host
        $GenerateReport = Read-Host "Do you want to export this report to Excel - Y/N?"
        If ($GenerateReport -eq "Y") {
            $date = Get-Date -Format "yyyy-MMM-dd-HHmmss"
            $xlsx_consolidation_output_file = "$Global:SnapshotManagement_home\$global:DefaultVIServer-SnapshotConsolidationReport-$Date.xlsx"     
            $snapshotConsolidationCollection | Export-Excel $xlsx_consolidation_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname Snapshots -AutoSize 
            Write-Host "`nReport generated : $xlsx_consolidation_output_file" -ForegroundColor Green

            $subj = "Snapshot consolidation report ..."
            Write-Host "Sending email report" -ForegroundColor Green
            Send-EmailReport $subj $xlsx_consolidation_output_file        
        }
    }
    
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu    
} # end function Get-VMsNeedingConsolidation

function Get-SnapshotAge ($targetAge) {
    #Clear-Host
    Write-Host "List all VMs with snapshots older than specified age." -ForegroundColor Green
    Write-Host "-----------------------------------------------------" -ForegroundColor Green

    $snapshotAgeCollection = @()

    foreach ($snap in Get-VM | Get-Snapshot) {
        $SnapshotAge = ((Get-Date) - $snap.Created).Days
        If ($SnapshotAge -gt $targetAge) {
            $snapAgeInfo = [PSCustomObject]@{
                "VM"                  = $snap.vm
                "Snapshot Name"       = $snap.name
                "Snapshot age (days)" = $SnapshotAge
            } # end $snapAgeInfo = [PSCustomObject]@
            $snapshotAgeCollection += $snapAgeInfo
        }
    }

    If ($snapshotAgeCollection.count -eq 0) {
        Write-Host "`nNo VMs found with snapshots older than $targetAge days found." -ForegroundColor Green
    } else {
        $snapshotAgeCollection | Sort-Object -Property "Snapshot age (days)" | Out-Host
    }

    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end function Get-SnapshotAge

function Get-SnapshotSize ($targetSize) {
    #Clear-Host 
    Write-Host "List all VMs with snapshots larger than specfied size." -ForegroundColor Green
    Write-Host "------------------------------------------------------" -ForegroundColor Green

    $snapshotSizeCollection = @()

    foreach ($snap in Get-VM | Get-Snapshot) {
        If (([math]::round($snap.sizeGB)) -gt $targetSize) {
            $snapSizeInfo = [PSCustomObject]@{
                "VM"                 = $snap.vm
                "Snapshot Name"      = $snap.name
                "Snapshot size (GB)" = [math]::round($snap.sizeGB)
            } # end $snapSizeInfo = [PSCustomObject]@
            $snapshotSizeCollection += $snapSizeInfo
        }
    }
    If ($snapshotSizeCollection.count -eq 0) {
        Write-Host "`nNo VMs with snapshots larger than $targetSize `bGB found." -ForegroundColor Green
    } else {
        $snapshotSizeCollection | Sort-Object -Property "Snapshot size (GB)" | Out-Host
    }
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end function Get-SnapshotSize

function Get-DatastoresWithSnapshot {
    #Clear-Host
    Write-Host "List all datastores and free space where VMs with snapshots reside." -ForegroundColor Green
    Write-Host "-------------------------------------------------------------------" -ForegroundColor Green
    $dsCollection = @()

    foreach ($snap in Get-VM -Location $dc | Get-Snapshot) {
        $ds = Get-Datastore -VM $snap.vm | Select-Object name, 
            @{n = "Capacity"; E = { [math]::round($_.CapacityGB) } }, 
            @{n = "FreeSpace"; E = { [math]::round($_.FreeSpaceGB) } }, 
            @{N = "PercentFree"; E = { [math]::round($_.FreeSpaceGB / $_.CapacityGB * 100) } } 

        $dsinfo = [PSCustomObject]@{
            "Datastore"           = $ds.name
            "FreeSpace GB"        = $ds.FreeSpace
            "Percentage free (%)" = $ds.PercentFree
        }
        $dsCollection += $dsinfo
    }
    $dsCollection | Sort-Object Datastore -unique | Out-Host
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu    
} # end function Get-DatastoreWithSnapshot

function Get-VMSnapshotEvents {
    $targetVM -eq $null

    #Clear-Host
    Write-Host "List all snapshot related tasks and events for a specific VM." -ForegroundColor Green
    Write-Host "-------------------------------------------------------------" -ForegroundColor Green
    $TargetVM = Read-Host "`nEnter the VM to check for snapshot related tasks and events."

    Check-VMExists $TargetVM

    Write-Host "`nLooking to retrieve any events related to snapshots for this VM" -ForegroundColor Green
    Write-Host "Note that message may have already rotated out of the vCenter logs." -ForegroundColor Green

    $VMSnapshotEvents = Get-VM $TargetVM | Get-VIEvent | Where-Object {$_.FullFormattedMessage -like "*snapshot*"} | 
        Select-Object FullFormattedMessage, Username, CreatedTime | Sort-Object -Property CreatedTime -Descending 
    If ($VMSnapshotEvents.count -eq 0) {
        Write-Host "`nNo snapshot related events found." -ForegroundColor Green
        Read-Host "`nPress ENTER to return to main menu."
        SnapshotManagementMenu
    }
    $VMSnapshotEvents | Out-Host | Format-Table -Autosize

    $GenerateReport = Read-Host "Do you want to export these messages to Excel - Y/N?"
    If ($GenerateReport -eq "Y") {
        $date = Get-Date -Format "yyyy-MMM-dd-HHmmss"
        $xlsx_snapshotevents_output_file = "$Global:SnapshotManagement_home\$global:DefaultVIServer-$targetVM-snapshotevents-$Date.xlsx"     
        $VMSnapshotEvents | Export-Excel $xlsx_snapshotevents_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "SnapshotEvents" -AutoSize 
        Write-Host "`nReport generated : $xlsx_snapshotevents_output_file" -ForegroundColor Green
    }
    Read-Host "`nPress ENTER to return to main menu."
    SnapshotManagementMenu

} # end function Get-VMSnapshotEvents
 
function Get-SnapshotPreCheck {
    Write-Host "`nPlaceholder" -ForegroundColor Green
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu    
} # end function Set-SnapshotPreCheck

# Based around https://enterpriseadmins.org/blog/scripting/get-vcenter-scheduled-tasks-with-powercli-part-1/
Function Get-AllScheduledSnapshots {
    #Clear-Host
    Write-Host "List ALL snapshot related scheduled tasks." -ForegroundColor Green
    Write-Host "------------------------------------------" -ForegroundColor Green

    If (Get-VIScheduledTasks | Where-Object { $_."Action Name" -eq 'CreateSnapshot_Task' }) {
        Write-Host "Snapshot scheduled tasks found (times in UTC)" -ForegroundColor Green
        Get-VIScheduledTasks | Where-Object { $_."Action Name" -eq 'CreateSnapshot_Task' } |
        Select-Object @{N = "VM"; E = { $_."Entity Name" } }, Name, "Next Run Time", Description, Notification, "Action Name" | Format-Table -autosize
    } else {
        Write-Host "`nNo snapshot tasks found." -ForegroundColor Green
    }
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu

} # end Function Get-AllScheduledSnapshots

# Based around https://enterpriseadmins.org/blog/scripting/get-vcenter-scheduled-tasks-with-powercli-part-1/
Function Get-ScheduledSnapshots {
    #Clear-Host
    Write-Host "List ALL snapshot related scheduled tasks still due to run." -ForegroundColor Green
    Write-Host "-----------------------------------------------------------" -ForegroundColor Green

    If (Get-VIScheduledTasks | Where-Object { $_."Action Name" -eq 'CreateSnapshot_Task' -AND $_."Next Run Time" -ne $null }) {
        Write-Host "Future snapshot scheduled tasks found (times in UTC)" -ForegroundColor Green
        Get-VIScheduledTasks | Where-Object { $_."Action Name" -eq 'CreateSnapshot_Task' -AND $_."Next Run Time" -ne $null } |
        Select-Object @{N = "VM"; E = { $_."Entity Name" } }, Name, "Next Run Time", Description, Notification, "Action Name" | Format-Table -autosize
    } else {
        Write-Host "`nNo future scheduled snapshot tasks found." -ForegroundColor Green
    }
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end Function Get-ScheduledSnapshots

# Based on https://communities.vmware.com/thread/541573
function New-ScheduleVMSnapshot {
    $targetVM -eq $null

    #Clear-Host
    Write-Host "Schedule a snapshot for a specific VM." -ForegroundColor Green
    Write-Host "--------------------------------------" -ForegroundColor Green

    # Gather the various bits of information that we're going to need.
    $targetVM = Read-Host "Enter the name of the VM as it appears in vSphere, that you wish to take a snapshot for"
    Check-VMExists $targetVM
    Write-Host "`nSet the time to schedule this for. It needs to be in a datetime format so " -ForegroundColor Green
    Write-Host "mm/dd/yyyy hh:mm and is in UTC (BST -1)" -ForegroundColor Green
    [Datetime]$targetHour = Read-Host "`nEnter the date and time to schedule the snapshot for."
    [Datetime]$SchedTime = $TargetHour.ToUniversalTime()
    $snapName = Read-Host "Enter the name for the snapshot"
    $snapDescription = Read-Host "Enter a description for the snapshot"
    $IncMem = Read-Host "Include memory - default is no - Y/N"
    If ($IncMem -eq "Y") {
        $snapMemory = $true
    } else {
        $snapMemory = $false
    }
    $Quiese = Read-Host "Quiesce the VM - default is no - Y/N"
    If ($Quiese -eq "Y") {
        $snapQuiesce = $true
    } else {
        $snapQuiesce = $false
    }

    $vm = Get-VM -Name $targetVM

    $si = get-view ServiceInstance
    $scheduledTaskManager = Get-View $si.Content.ScheduledTaskManager

    $spec = New-Object VMware.Vim.ScheduledTaskSpec
    $spec.Name = "Snapshot $($vm.name) at $Schedtime"
    $spec.Description = "Take a snapshot of $($vm.Name)"
    $spec.Enabled = $true
#    $spec.Notification = $emailAddr
    $spec.Scheduler = New-Object VMware.Vim.OnceTaskScheduler
    $spec.Scheduler.runat = $SchedTime
    $spec.Action = New-Object VMware.Vim.MethodAction
    $spec.Action.Name = "CreateSnapshot_Task"

    @($snapName,$snapDescription,$snapMemory,$snapQuiesce) | ForEach-Object{
        $arg = New-Object VMware.Vim.MethodActionArgument
        $arg.Value = $_
        $spec.Action.Argument += $arg
    }

    Try {
        $scheduledTaskManager.CreateObjectScheduledTask($vm.ExtensionData.MoRef, $spec)
    } Catch {
        Write-Host "`nEncountered an error when trying to create the scheduled task." -ForegroundColor Green
        Write-Host "Please review the events in vSphere." -ForegroundColor Green
    }
    Read-Host "`nPress ENTER to return to main menu."
    SnapshotManagementMenu
}

function New-VMSnap {
    $targetVM -eq $null

    #Clear-Host
    Write-Host "Take a snapshot for a specific VM." -ForegroundColor Green
    Write-Host "----------------------------------" -ForegroundColor Green

    $targetVM = Read-Host "`nEnter the name of the VM to take a snapshot of" 
    Check-VMExists $targetVM
    $SnapName = Read-Host "Enter a name for the snapshot"
    $SnapDesc = Read-Host "Enter a description"
    $IncludeMemory = Read-Host "Include memory - Y/N?"
    $Quiesce = Read-Host "Quiesce the VM - Y/N?"
    
    # ugly ... also, need to check against powerstate when selecting Yes for quiesce and memory inclusion
    # If the VM isn't powered on, we can't include those
    Write-Host "It is assumed you have run the pre-check for this VM, and so know there's sufficient free"
    Write-Host "space on the datastore to take the VM. If not, you should run that option first."
    $ContinuewithTakingSnapshot = Read-Host "Do you wish to continue to take the snapshot : Y/N?"
    If ($ContinuewithTakingSnapshot -eq "Y") {
        Write-Host "Taking the snapshot"
        Try {
            If ($IncludeMemory -eq "Y" -AND $Quiesce -eq "Y") {
                If ($targetVM.powerstate -eq "PoweredOff") {
                    Write-Host "The VM is powered off, so cannot include memory and quiescing the VM for a snapshot."
                    Write-Host "Returning to main menu."
                }
                Get-VM $targetVM | New-Snapshot -Name $SnapName -Description $SnapDesc -Memory -Quiesce -ErrorAction Stop

            } elseif ($IncludeMemory -eq "Y") {
                If ($targetVM.powerstate -eq "PoweredOff") {
                    Write-Host "The VM is powered off, so take snapshot with memory."
                    Write-Host "Returning to main menu."
                }
                Get-VM $targetVM | New-Snapshot -Name $SnapName -Description $SnapDesc -Memory -ErrorAction Stop

            } elseif ($Quiesce -eq "Y") {
                If ($targetVM.powerstate -eq "PoweredOff") {
                    Write-Host "The VM is powered off, so cannot quiesce the VM."
                    Write-Host "Returning to main menu."
                }
                Get-VM $targetVM | New-Snapshot -Name $SnapName -Description $SnapDesc -Quiesce -ErrorAction Stop

            } else {
                Get-VM $targetVM | New-Snapshot -Name $SnapName -Description $SnapDesc -ErrorAction Stop 
            }

            $VMSnapshotEvents = Get-VM $TargetVM | Get-VIEvent -MaxSamples 10 | Where-Object {$_.FullFormattedMessage -like "*snapshot*"} | 
            Select-Object FullFormattedMessage, Username, CreatedTime | Sort-Object -Property CreatedTime -Descending 
            Write-Host "`n10 most recent snapshot tasks and events - please review to confirm snapshot you are seeing a successful snapshot message." -ForegroundColor Green
            $VMSnapshotEvents | Out-Host #| Format-Table -AutoSize
        } Catch {
            Write-Host "There was an error in taking the snapshot."
            Write-Host "Please logon to the VC and check the logs"
            Read-Host "Press ENTER to return to main menu."
            SnapshotManagementMenu
        }
    } else {
        Read-Host "Press ENTER to return to main menu."
        SnapshotManagementMenu    
    }
    Write-Host "`nSnapshot taken" -ForegroundColor Green
    # Let's print the details of the snapshot as evidence
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end function New-VMSnap

function New-MultiVMSnap {
    #Clear-Host
    Write-Host "Take a snapshot for multiple VMs." -ForegroundColor Green
    Write-Host "---------------------------------" -ForegroundColor Green

$MultVMSnapBlurb = @"

This option will take snapshots on a number of VMs.
Things to consider :
1) VMs are to be in a .csv file, with a single column named "vmname".
2) The snapshot name, description, inclusion of memory and quiescing options will be the same for all VMs in this list.
In other words, if you choose to include memory in the snapshot, each of the VMs in the list will have their snapshot taken 
this way.

It is assumed you have completed pre-requisite checks and that you know there is sufficient free space on the underlying
datastores. If you haven't, it's recommended you do this first!

Proceed now with supplying the name of the .csv file containing all the VMs to be snapshot.
"@

    $MultVMSnapBlurb

    # Figure out and try to read in the .csv file
    $srcFile = Read-Host "`n Enter the source .csv file."

    # Check if the file exists. If not, offer to return to menu or quit.
    If (-not (Test-Path -path $srcFile)) {
        Write-Host "File doesn't exist." -ForegroundColor Red
        Read-Host "Press ENTER to return to main menu."
        SnapshotManagementMenu
    }

    $vmsToSnap = Import-Csv $srcFile

    # Ask questions applicable to ALL the VMs - inc memory, quiesce, time
    $SnapshotName = Read-Host "`nEnter a name for the snapshot - this will be common to all the snapshots taken."
    $SnapDesc = Read-Host "Enter a generic description for the snapshot"
    $IncludeMemory = Read-Host "Include memory - Y/N?"
    $Quiesce = Read-Host "Quiesce the VM - Y/N?"

    Write-Host "Taking the snapshot"
    ForEach ($vm in $vmsToSnap) {
        $CurrentVM = Get-VM $VM.vmname -ErrorAction SilentlyContinue
        Try {
            If ($IncludeMemory -eq "Y" -AND $Quiesce -eq "Y") {
                If ($CurrentVM.powerstate -eq "PoweredOff") {
                    Write-Host "The VM is powered off, so cannot include memory and quiescing the VM for a snapshot."
                    Write-Host "Returning to main menu."
                }
                $CurrentVM | New-Snapshot -Name $SnapshotName -Description $SnapDesc -Memory -Quiesce -ErrorAction Stop

            } elseif ($IncludeMemory -eq "Y") {
                If ($CurrentVMVM.powerstate -eq "PoweredOff") {
                    Write-Host "The VM is powered off, so take snapshot with memory."
                    Write-Host "Returning to main menu."
                }
                $CurrentVM | New-Snapshot -Name $SnapshotName -Description $SnapDesc -Memory -ErrorAction Stop

            } elseif ($Quiesce -eq "Y") {
                If ($CurrentVM.powerstate -eq "PoweredOff") {
                    Write-Host "The VM is powered off, so cannot quiesce the VM."
                    Write-Host "Returning to main menu."
                }
                $CurrentVM | New-Snapshot -Name $SnapshotName -Description $SnapDesc -Quiesce -ErrorAction Stop

            } else {
                Write-Host "Taking snapshot for $CurrentVM"
                Get-VM $CurrentVM | New-Snapshot -Name $SnapshotName -Description $SnapDesc -ErrorAction Stop 
            }
        } Catch {
            Write-Host "There was an error in taking the snapshot."
            Write-Host "Please logon to the VC and check the logs"
        }
    } # end foreach ($vm in $vmsToSnap)

    Write-Host "`nSnapshot taken" -ForegroundColor Green
    # Let's print the details of the snapshot as evidence
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu


} # end function New-MultiVMSnap
 
function Remove-VMSnapshot {
    $null -eq $targetVM

    ##Clear-Host 
    Write-Host "Remove snapshot(s) for a specific VM." -ForegroundColor Green
    Write-Host "-------------------------------------" -ForegroundColor Green

    $targetVM = Read-Host "Enter the VM name to remove snapshots from"
    Check-VMExists $TargetVM
    # Print the current list of snapshot
    $listOfSnaps = get-vm $targetVM | get-snapshot | Select-Object VM, Name, IsCurrent, Created, Parent, Children, 
        @{n="Snapshot size GB"; E= { [math]::round($_.sizeGB) } } | Format-Table -Autosize
    If ($listOfSnaps.count -eq 0) {
        Write-Host "`nVM has no snapshots" -ForegroundColor Green
        Read-Host "Press ENTER to return to main menu."
        SnapshotManagementMenu    
    } else {
        $listOfSnaps | Out-Host

        $DeleteAllSnaps = Read-Host "Do you wish to delete ALL snapshots on this VM - Y/N?"
        If ($DeleteAllSnaps -eq "Y") {
            get-vm $targetVM | get-snapshot| Remove-Snapshot -confirm:$False
            $listOfSnaps = get-vm $targetVM | get-snapshot | 
                Select-Object VM, Name, IsCurrent, Created, Parent, Children, 
                @{n="Snapshot size GB"; E= { [math]::round($_.sizeGB) } } #| Format-Table -Autosize

            # Print the 10 most recent snapshot events
            $VMSnapshotEvents = Get-VM $TargetVM | Get-VIEvent -MaxSamples 10 | Where-Object {$_.FullFormattedMessage -like "*snapshot*"} | 
            Select-Object FullFormattedMessage, Username, CreatedTime | Sort-Object -Property CreatedTime -Descending 
            Write-Host "`n10 most recent snapshot tasks and events - please review to confirm snapshot you are seeing a successful snapshot deleted message." -ForegroundColor Green
            $VMSnapshotEvents | Format-Table -AutoSize

            Write-Host "`nUpdated list of snapshots - should be nothing returned." -ForegroundColor Green
            $listOfSnaps | Out-Host | Format-Table -AutoSize

            Write-Host "No snapshot with that name." -ForegroundColor Green
            Read-Host "Press ENTER to return to main menu."
            SnapshotManagementMenu            

        } else {
                $SnapToDelete = Read-Host "Enter the name of the snapshot to delete"
                $DeleteChildren = Read-Host "Delete any child snapshots - Y/N?"
                If (Get-VM $targetVM | Get-Snapshot -Name $snapToDelete) {
                    If ($DeleteChildren -eq "Y") {
                        Get-VM $targetVM | Get-Snapshot -Name $snapToDelete | Remove-Snapshot -RemoveChildren -confirm:$false

                        # Print the 10 most recent snapshot events
                        $VMSnapshotEvents = Get-VM $TargetVM | Get-VIEvent -MaxSamples 10 | Where-Object {$_.FullFormattedMessage -like "*snapshot*"} | 
                        Select-Object FullFormattedMessage, Username, CreatedTime | Sort-Object -Property CreatedTime -Descending 
                        Write-Host "`n10 most recent snapshot tasks and events - please review to confirm snapshot you are seeing a successful snapshot deleted message." -ForegroundColor Green
                        $VMSnapshotEvents | Format-Table -AutoSize
                    
                        $null -eq $listOfSnaps
                        $listOfSnaps = get-vm $targetVM | get-snapshot | 
                            Select-Object VM, Name, IsCurrent, Created, Parent, Children, 
                            @{n="Snapshot size GB"; E= { [math]::round($_.sizeGB) } } | Format-Table -Autosize
                        Write-Host "`nUpdated list of snapshots." -ForegroundColor Green
                        $listOfSnaps | Out-Host #| Format-Table -AutoSize

                        Read-Host "Press ENTER to return to main menu."
                        SnapshotManagementMenu            
            
                    } else {
                        Get-VM $targetVM | Get-Snapshot -Name $snapToDelete | Remove-Snapshot -confirm:$false

                        # Print the 10 most recent snapshot events
                        $VMSnapshotEvents = Get-VM $TargetVM | Get-VIEvent -MaxSamples 10 | Where-Object {$_.FullFormattedMessage -like "*snapshot*"} | 
                            Select-Object FullFormattedMessage, Username, CreatedTime | Sort-Object -Property CreatedTime -Descending 
                        Write-Host "`n10 most recent snapshot tasks and events - please review to confirm snapshot you are seeing a successful snapshot deleted message." -ForegroundColor Green
                        $VMSnapshotEvents | Format-Table -AutoSize

                        $listOfSnaps = get-vm $targetVM | get-snapshot | 
                            Select-Object VM, Name, IsCurrent, Created, Parent, Children, 
                            @{n="Snapshot size GB"; E= { [math]::round($_.sizeGB) } } #| Format-Table -Autosize
                        Write-Host "`nUpdated list of snapshots." -ForegroundColor Green
                        $listOfSnaps | Format-Table -AutoSize

                        Read-Host "Press ENTER to return to main menu."
                        SnapshotManagementMenu            
            
                    } # end if else ($DeleteChildren -eq "Y")
                } else {
                    Write-Host "No snapshot with that name." -ForegroundColor Green
                    Read-Host "Press ENTER to return to main menu."
                    SnapshotManagementMenu            
                } # end if else (Get-VM $targetVM | Get-Snapshot -Name $SnapToDelete)
        } # end if  else If ($DeleteAllSnaps -eq "Y")

    } # end else

} # end function Remove-VMSnapshot

function Revert-VMSnapshot {
    $targetVM -eq $null

    #Clear-Host 
    Write-Host "Revert snapshot for a specific VM." -ForegroundColor Green
    Write-Host "----------------------------------" -ForegroundColor Green

    $targetVM = Read-Host "Enter the VM name to revert snapshot on."
    Check-VMExists $TargetVM
    # Print the current list of snapshot
    $listOfSnaps = get-vm $targetVM | get-snapshot | 
        Select-Object VM, Name, IsCurrent, Created, Parent, Children, @{n="Snapshot size GB"; E= { [math]::round($_.sizeGB) } } | Format-Table -Autosize
    If ($listOfSnaps.count -eq 0) {
        Write-Host "`nVM has no snapshots" -ForegroundColor Green
        Read-Host "Press ENTER to return to main menu."
        SnapshotManagementMenu    
    } else {
        $listOfSnaps | Out-Host
        $SnapToRevert = Read-Host "Enter the name of the snapshot to revert to."
        If (Get-VM $targetVM | Get-Snapshot -Name $snapToRevert) {
            Write-Host "Reverting to $SnapToRevert" -ForegroundColor Green
            Get-VM $targetVM | Set-VM -Snapshot $SnapToRevert -Confirm:$false
            $listOfSnaps = get-vm $targetVM | get-snapshot | 
                Select-Object VM, Name, IsCurrent, Created, Parent, Children, @{n="Snapshot size GB"; E= { [math]::round($_.sizeGB) } } | Format-Table -Autosize

            # Print the 10 most recent snapshot events
            $VMSnapshotEvents = Get-VM $TargetVM | Get-VIEvent -MaxSamples 10 | Where-Object {$_.FullFormattedMessage -like "*snapshot*"} | 
                Select-Object FullFormattedMessage, Username, CreatedTime | Sort-Object -Property CreatedTime -Descending 
            Write-Host "`n10 most recent snapshot tasks and events - please review to confirm snapshot you are seeing a successful snapshot revert message." -ForegroundColor Green
            $VMSnapshotEvents | Format-Table -AutoSize

            Write-Host "`nUpdated snapshot list :" -ForegroundColor Green
            $listOfSnaps | Out-Host
        } else {
            Write-Host "No snapshot with that name" -ForegroundColor Green
            Read-Host "Press ENTER to return to main menu."
            SnapshotManagementMenu        
        }
    }
    Read-Host "Press ENTER to return to main menu."
    SnapshotManagementMenu
} # end function Revert-VMSnapshot

## Menu
function SnapshotManagementMenu {

    $Menu = @"
`t`t`tSnapshot management
`t`t`t-------------------
Reporting tasks
---------------
1) Simple list of all VMs with snapshots.
2) Detailed list all snapshots in the environment.
3) Check an individual VM for existing snapshots.
4) List all VMs with snapshots requiring consolidation.
5) List all VMs with snapshots older than specified number of days.
6) List all VMs with snapshots larger than specified size in GB.
7) Report free space on datastores with VMs with snapshots.
8) List task and events messages related to snapshots for specific VM.

Pre-snapshot checks
-------------------
9) Perform pre-checks for taking a snapshot - single VM only.

Scheduled task related
----------------------
10) List all snapshot scheduled tasks in the environment.
11) List only the snapshot scheduled tasks still due to run.

Taking snapshots
----------------
12) Schedule a snapshot for a VM.
13) Take a snapshot for a single VM now.
14) Take snapshots on multiple VMs now.

Deleting or reverting snapshots
-------------------------------
15) Delete snapshot(s) from a single VM.
16) Revert a snapshot.

Press Q to quit.

"@

    Clear-Host
    $Menu
    $SnapshotMenuChoice = $null
    $SnapshotMenuChoice = Read-Host "`nEnter your option."

    Switch ($SnapshotMenuChoice) {
        1 {
            Get-ListOfVMsWithSnapshots
        }

        2 {
            Get-SnapshotReport
        }   

        3 {
            Check-VMForSnapshot
        }

        4 {
            Get-VMsNeedingConsolidation    
        }

        5 {
            Write-Host "`nThis will check for all VMs with a snapshot older than the number of days you specify." -ForegroundColor Green
            Write-Host "Those VMs snapshots will probably warrant closer inspection.`n" -ForegroundColor Green
            $Age = Read-Host "`nEnter the number of days."
            Get-SnapshotAge $Age
        }

        6 {
            Write-Host "`nThis will check for all VMs with a snapshot larger in size (GB) than the value you specify." -ForegroundColor Green
            Write-Host "Those VMs snapshots will probably warrant closer inspection.`n" -ForegroundColor Green
            Write-Host "`nNote that the snapshot size will be rounded to the closest GB." -ForegroundColor Green
            $Size = Read-Host "`nEnter the size (GB) to check against."
            Get-SnapshotSize $Size
        }

        7 {
            Get-DatastoresWithSnapshot
        }

        8 {
            Get-VMSnapshotEvents
        }

        9 {
            Get-SnapshotPreCheck
        }

        10 {
            Get-AllScheduledSnapshots
        }

        11 {
            Get-ScheduledSnapshots
        }

        12 {
            New-ScheduleVMSnapshot
        }

        13 {
            Clear-Host
            New-VMSnap
        }

        14 {
            New-MultiVMSnap
        }

        15 {
            #Clear-Host
            Remove-VMSnapshot
        }

        16 {
            Revert-VMSnapshot
        }

        q {
            Write-Host "Quitting."
            Break
        }
        default {
            Write-Host "Invalid option. Quitting."
            Break
        }
    } # end Switch

} # end function SnapshotManagementMenu

## Main function to run
function SnapshotManagement {
    $banner = @"
SSSSS                  p ppp           h                                                                ggg g
S     S                 pp   p          h                 t                                             g   gg                                    t
S                       p    p          h                 t                                             g    g                                    t
S       n nnn    aaaa   p    p   ssss   h hhh    oooo   ttttt           m m mm   aaaa   n nnn    aaaa   g    g   eeee   m m mm   eeee   n nnn   ttttt
 SSSSS  nn   n       a  pp   p  s    s  hh   h  o    o    t             mm m  m      a  nn   n       a  g   gg  e    e  mm m  m e    e  nn   n    t
      S n    n   aaaaa  p ppp    ss     h    h  o    o    t             m  m  m  aaaaa  n    n   aaaaa   ggg g  eeeeee  m  m  m eeeeee  n    n    t
      S n    n  a     a p          ss   h    h  o    o    t             m  m  m a     a n    n  a     a      g  e       m  m  m e       n    n    t
S     S n    n  a    aa p       s    s  h    h  o    o    t  t          m  m  m a    aa n    n  a    aa g    g  e    e  m  m  m e    e  n    n    t  t
 SSSSS  n    n   aaaa a p        ssss   h    h   oooo      tt           m  m  m  aaaa a n    n   aaaa a  gggg    eeee   m  m  m  eeee   n    n     tt
"@
    Clear-Host
    #$banner
    Set-InitializeScript
#    Clear-Host
    Get-VCConnectionState
    SnapshotManagementMenu
} # end function SnapshotManagement       