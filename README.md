# powershell
After a lot of late nights patching critical servers until past 2AM, I thought that there must be a better way than what I was doing. I ended up developing a set of scripts that essentially run Windows MSU files in parallel on list of machines. After a lot of refining and testing, I ran it on the last Critical Patch night and saved several hours of time.




This script assumes that:

You have this folder created: C:\WindowsUpdates

You have both UpdateEndpoints.ps1 and UpdateEndpoints-RemoteMachine.ps1 located in the C:\WindowsUpdates folder

When you download an update from the Microsoft Catalog, you will keep the KBXXXXXX in the name. It can be just named for the KB, or you can copy and paste the entire update title, whatever you want - just make sure the full KB is in the name of the file, as this is how the script tracks what each server needs

You have your list of machines in CSV format that the script can access. The CSV should be one column, with endpoints in the first cell, followed by all the machines you want to update

This is a set of two scripts; UpdateEndpoints.ps1 and UpdateEndpoints-Local.ps1. The reason I broke this up into two separate files was to show the progress of installing a CAB/MSU file, which from all my testing was only possible to show in a native PowerShell session. It's possible to run the same script through Invoke-Command on all these servers at once, but the major drawback is the fact that you will not see any of the progress, and so you won't know the status of the update, how many updates are left, if it's hung up, etc. There may be a better way to do this that a PowerShell professional would be able to pull off, but this is the best way I found.

Essentially, this script does everything you would do manually and attempts to take the pain and overall time consumption out of the process. This is what the script does:

Load the list of machines from the CSV you provide

Invoke the WSUS service to check for any uninstalled updates

Displays on the console what each server is missing

When all servers have been check, displays a list of all unique missing updates (and optionally emails you a summary as well)

Scans the C:\WindowsUpdates folder for these updates, and displays what is present and what still may need to be downloaded. Note: it does not require all patches to be downloaded, it will only install what you decided to download

Next it will loop through each machine, check what is missing for that particular machine, and if that exists locally, it will copy it to the remote machine (creating the C:\WindowsUpdates folder is it does not already exist). It will also copy over the UpdateEndpoints-RemoteEndpoints.ps1 file to that folder.

It then opens a new PowerShell window for each remote machine and enters a PSSession connecting to it. In this window will be printed the command that you will need to paste into the session once it's connected. This will run UpdateEndpoints-RemoteEndpoints.ps1 on the remote machine. I tried making this automatic, but couldn't figure out how to do this as it never properly waited for the remote session to be established and would run in the context of my local machine.

At this point each remote machine will be extracting each MSU in the C:\WindowsUpdates folder into its own folder, and then installing the relevant CAB file via the PowerShell cmdlet Add-WindowsPackage. Once it has looped through each MSU, if it the Add-WindowsPackage cmdlet had errors you will see it on the console. The script will also check for local WUSA errors in the event log and report any findings. If not are found, it will print so to the console and prompt for a reboot of the server. Note: While this script will copy over all missing updates you have manually downloaded, it will only attempt to install the MSU files. Any other files ending is EXE or otherwise will be copied over for easier manual installation, but will not have its installation attempted via this script, as there are too many variables to consider. Further fine-tuning of this script may allow for this in the future.

Once all the remote PowerShell sessions started, the original PowerShell window begins checking the connection to the remote machines and prints any changes to the console. This is not perfect, but serves to show you when a remote server has rebooted so you can remote in and check it out to make sure the update took.

That's about it. This could undoubtedly be refined further to give more useful information to the end user, be able to also apply other EXE updates such as the Malicious Software Removal tool, SQL CU updates, etc. The purpose of this script is to save time and mental effort necessary to update 13+ critical servers all at the same time late on a Saturday night, so if it fails to save you time and effort, don't use it.
