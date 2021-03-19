$updateFolder = "C:\WindowsUpdates\*";
$FileTime = Get-Date -format 'yyyy.MM.dd-HH.mm'
$updates = Get-Childitem $updateFolder -Include "*.msu";
$Qty = $updates.count
Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -first 5 # Print currently installed updates

$updates | % {
    $KB = $_.Name.split(".")[0]
    if (!(Test-Path "c:\WindowsUpdates\$($KB)-extract")){
        New-Item -ItemType directory -Path "c:\WindowsUpdates\$($KB)-extract"
    }
    #cmd /c "winrs.exe -r:$($env:COMPUTERNAME) wusa.exe '$($_.FullName)' /extract:'C:\WindowsUpdates\extracts'"
    expand -f:* $_.FullName "C:\WindowsUpdates\$($KB)-extract"
    $cabs = Get-Childitem "C:\WindowsUpdates\$($KB)-extract\*" -Include "*.cab";
    ForEach ($cab IN $cabs) {
        if (!($cab.Name -match "WSUSSCAN")) {
            Write-Host "Starting Update Cab $($cab.FullName) on $($env:COMPUTERNAME)";
            #cmd /c "winrs.exe -r:$($env:COMPUTERNAME) dism.exe /online /add-package /PackagePath:$($cab.FullName)"
            #cmd /c "dism.exe /online /add-package /PackagePath:$($cab.FullName)"
            #DISM /online /NoRestart /Add-Package /PackagePath:"$($cab.FullName)" | Tee-Object -Variable cmdOutput # Note how the var name is NOT $-prefixed
            Add-WindowsPackage -Online -NoRestart -PackagePath "$($cab.FullName)"
            Write-Host "Finished Update Cab $($cab.FullName) on $($env:COMPUTERNAME)";
        }
    }
}
Write-Host "Finished Update $($_) on $($env:COMPUTERNAME)";
if (!(Test-Path C:\WindowsUpdates\logs)) {
    New-Item -ItemType Directory -Path C:\WindowsUpdates\logs
}
if (Test-Path c:\WindowsUpdates\Wusa.log){
    Move-Item -Path "c:\WindowsUpdates\Wusa.log" -Destination "c:\WindowsUpdates\logs\Wusa.$($_.Name).$($FileTime).evtx"
}
$Qty = --$Qty
$Qty
if ($Qty -eq 0){

    # Check the Event logs for udate errors
    $StartTime = [DateTime]::Now.AddHours(-8);
    $EndTime = [DateTime]::Now
    $errorLogs = Get-WinEvent -FilterHashTable @{
        LogName='Setup'
        Level=2
        ProviderName='Microsoft-Windows-WUSA'
        StartTime=$StartTime
        EndTime=$EndTime
    } -erroraction 'silentlycontinue'
    if ($errorLogs.count -gt 0) {
        # Errors were found. Break the process and manually reconcile.
        Write-Host "WUSA Event Log error found on $($env:COMPUTERNAME):"
        $errorLogs
        break;
    }
    else {
        # Cleanup the update files
        Get-ChildItem -path C:\WindowsUpdates -Directory -recurse | % {
            if ($_.Name -match "extract") {
                Remove-Item -LiteralPath "$($_.FullName)" -Force -Recurse
            }
        }
        Get-ChildItem *.msu | foreach { Remove-Item -Path $_.FullName }
        Get-ChildItem *.exe | foreach { Remove-Item -Path $_.FullName }
        Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -first 5 # Print currently installed updates
        $confirmation = Read-Host "All updates for $($env:COMPUTERNAME) have completed with no errors found. Restart $($env:COMPUTERNAME) y/n?"
        if ($confirmation -eq 'y') {
            Write-Host "Restarting $($env:COMPUTERNAME)...";
            #Restart-Computer
            restart-computer -force
        } else {
            Write-Host "You will need to manually restart $($env:COMPUTERNAME)";
        }
    }
}