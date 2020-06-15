# Snapshot Related Tasks

This script is a collection of functions of some of the common tasks and information that I like to have available when dealing with snapshots. It is still a work in progress. As is always the case. test first in a non production environment, and use at your own risk.

It's born out of working on a lot of different environments where, unfortunately, quite often management of snapshots is, frankly, not very good.

To run the script, simply dot source it and run the SnapshotManagement function.

    . .\SnapshotManagement.ps1
    SnapshotManagement

If looking to send emails, fill in the values in the Send-EmailReport function, otherwise comment out any references to that function to avoid errors.

It's split into a few sections :

## Reporting
A few options to get various bits of information on the status of VMs with snapshots, ages, size and the status of the underlying datastores used by the VMs.

## Pre-snapshot checks
Environments I work on often have very little free space on the datastores in general, so this is just about checking what the status is prior to taking a snapshot - you will need to decide if the free space available works for you or not.

## Scheduled snapshots
Lists the snapshot based scheduled tasks in vCenter (both ALL tasks and only those due to run). The code for these functions is primarily reworked from [1] in the Acknowledgements.

## Taking snapshots
Options to take snapshots now, for both a single VM and multiple VMs. Multiple VMs relies on a source .csv file, with a column labelled "vmname", and the names of the VMs you wish to snap. Snapshot options - ie, name, description, memory, quiesce etc, are basically then going to be the same for ALL the VMs in the .csv

There is also the option to create a scheduled task to take a snapshot on a single VM. Be very wary of timezones on this. The code for this function is primarily reworked from [3] in the Acknowledgements. 

## Deleting and reverting snapshots
Options to delete single or ALL snapshots on a VM, and also to revert to a specific snapshot.

## Acknowledgments
The script uses, either directly or (lightly) modified, code from the following :
1. https://enterpriseadmins.org/blog/scripting/get-vcenter-scheduled-tasks-with-powercli-part-1/ - the Get-VIScheduledTasks code is leveraged in a few of the functions.
2. https://communities.vmware.com/thread/541573 - specifically Luc Dekens example of scheduling a snapshot.

3. In addition, the ability to generate Excel based reports is based on use of the ImportExcel module available at https://www.powershellgallery.com/packages/ImportExcel/7.1.0

## Known issues
- Having problems if clearing the console when choosing each option - console isn't properly clearing for some functions - uncertain why.
- If ImportExcel isn't installed, currently the script doesn't offer to import it.
- There's quite a bit of code repetition, so a few additional functions probably needed to handle that.