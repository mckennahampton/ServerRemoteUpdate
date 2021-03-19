# Example usage:
# >. .\UpdateEndpoints.ps1
# >updateEndpoints -endpointCSV 'C:\path\to\endpoints.csv' -emailSummary 'y' -toEmail 'example@place.org' -fromEmail 'example@place.org' -smtpServer 'smtp.test.org'
#
# $toEmail - email to receive the email summary
# $fromEmail - send-as for the summary email
# $smtpServer - SMTP server for email summary
# $emailSummary - y/n
# $endpointCSV - CSV location of machines to update
#
# It is expected that the first row in the CSV is "endpoints", followed by each endpoint you want to update,
# ie endpoints,test-server1,test-server2,test-server3
#
# This file is used in conjunction with the UpdateEndpoints-Local.ps1 file, which is expected in C:\WindowsUpdates.
# If that file is not present, you will be alerted of this. That file is necessary, as it is compied over to the
# remote endpoint and ran in a remote Powershell session, so as to allow you to see the progress for each CAB file
# installation, and any error messages that may come from it. After trying everything I could to get this to work
# with Invoke-Command for all these servers at once, I never found a way to show the progress of the installation
# of the CAB files in through that manner.
#
#
function updateEndpoints( $endpointCSV, $emailSummary, $toEmail, $fromEmail, $smtpServer) {
    $test = Test-Path -path "C:\WindowsUpdates";
    If ($test -eq $true) {} else {
        Write-Host "Local update folder does not exist; creating C:\WindowsUpdates..."
        New-Item -ItemType directory -Path "C:\WindowsUpdates"
    }
    $test = Test-Path -path "C:\WindowsUpdates\logs";
    If ($test -eq $true) {} else {
        Write-Host "Local update log folder does not exist; creating C:\WindowsUpdates\logs..."
        New-Item -ItemType directory -Path "C:\WindowsUpdates\logs"
    }
    $LogPathName = "C:\WindowsUpdates\logs\patching-$(Get-Date -Format 'yyyy.MM.dd-HH.mm').log"
    Start-Transcript $LogPathName
    $serverList = @();
    Import-Csv $endpointCSV | Foreach-Object {$serverList += @($_.endpoints)};
    $serverMasterList = @();
    $missingUpdatesMasterList = @();
    ForEach ($server IN $serverList)
    {
        # First check if the update service is running
        $wus = Invoke-Command -Computer $server -ScriptBlock {
            Get-Service -Name wuauserv
        }
        $bits = Invoke-Command -Computer $server -ScriptBlock {
            Get-Service -Name bits
        }
        $origStatus=$wus.Status
        $origStartupType=$wus.StartType
        Write-Host "$($server)";
        Write-Host "Update service is in the $($origStatus) state and its startup type is $($origStartupType)" -verbose;
        if($origStatus -eq "Stopped"){

            Write-Host "Starting windows update service" -verbose
            Start-Service -Name wuauserv
        }
        if($bits.status -eq "Stopped") {
            Start-Service -Name bits
        }
        $emailUpdateSection = ""; # Reset the EmailUpdateString
        $OS = Get-WmiObject -Computer $server -Class Win32_OperatingSystem # Get the OS description
        $message = "$($server) - $($OS.Caption)" # Write server and OS to host
        $underlineChar = '-';
        $uLine = $underlineChar * $message.length;
        write-host -Object $message;
        write-host -Object $uLine;
        $objSession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$server)); # Create remote sessions to check for missing updates
        $objSearcher = $objSession.CreateUpdateSearcher(); # Creat the search object
        $Result = $objSearcher.Search("IsInstalled=0"); # Search for updates with the given citeria
        $updateTitles = $Result.Updates | Format-Table -property Title # List the updates in easily readable format
        $updateKb = @();
        $missingUpdateKB = @();
        $Result.Updates | % {
            if (($_.Title -ne "Microsoft driver update for Generic / Text Only") -and (-not($_.Title.Contains("Windows Malicious Software Removal Tool")))) {
                Write-Host $_.Title; # Print to console
                $emailUpdateSection = $emailUpdateSection  + "$($_.Title)<br>" # Add HTML section for email
                -split $_.Title  | % {if (($_ -like "KB*") -or ($_ -like "(KB*")) {
                    $updateKB = $_;
                    If (($updateKB[0] -eq "(") -and ($updateKB[-1] -eq ")")) {
                        $updateKB = $updateKB.Replace("(", "");
                        $updateKB = $updateKB.Replace(")", "");
                    }
                    $missingUpdateKB += $updateKB;
                    }
                } # Find the KB number for this update for file matching
                if ($missingUpdatesMasterList.Title -notcontains $_.Title) { # Add any new missing updates to the master missing update list
                    $missingUpdatesMasterList += [pscustomobject]@{
                        Title = $_.Title;
                        KB = $updateKB
                    };
                }
            }
        }
        $emailUpdateSection = $emailUpdateSection  + "<br><br>"
        $thisServer = [pscustomobject]@{
            ServerName = $server;
            OS = $OS.caption;
            MissingUpdates = $updateTitles;
            MissingKB = $missingUpdateKB;
            EmailUpdateString = $emailUpdateSection
        };
        $serverMasterList += $thisServer; # Add the server object to the array
        Write-Host "`n"
    }
    $emailBody = "";
    $serverMasterList | % {
        $emailBody = $emailBody + "<b>$($_.ServerName)</b> - $($_.OS)<br><hr>"; # Add server to email HTML body
        $emailBody = $emailBody + $_.EmailUpdateString; # Add missing updates to email HTML body
    }
    $emailBody = $emailBody + "Updates Needed:<br><hr>";
    $missingUpdatesMasterList | % {
        $emailBody = $emailBody + $_.Title + "<br>";
    }
    if (($emailSummary -ne $null) -or ($emailSummary -ne 'n')) {
        Write-Host "Email summary is set to be sent, checking for to-email, from-email, and SMTP server...";
        if (($toEmail -ne $null) -and ($fromEmail -ne $null) -and ($smtpServer -ne $null)) {
            Write-Host "Values supplied for to-email, from-email, and SMTP server. Attempting to send email summary...";
            send-mailmessage -to $toEmail -from $fromEmail -subject "Missing updates on Critical Servers" -smtpserver $smtpServer -BodyAsHtml $emailBody
        } else {
            Write-Host "One of the below values is missing, and thus the summary email cannot be sent:`n";
            Write-Host "To-email is set to $($toEmail)`n";
            Write-Host "From-email is set to $($fromEmail)`n";
            Write-Host "SMTP Server is set to $($smtpServer)`n";
        }
    }
    Write-Host "Checking downloaded updates from C:\WindowsUpdates...";
    $doneCheckingLocal = $false
    DO {
        # Check if UpdateEndpoints-Local.ps1 exists where expected
        $test = Test-Path -path "C:\WindowsUpdates\UpdateEndpoints-Local.ps1";
        If ($test -eq $true) {
            Write-Host "C:\WindowsUpdates\UpdateEndpoints-Local.ps1 is present." -ForegroundColor Green
            $doneCheckingLocal = $true
        } else {
            $response = Read-Host "C:\WindowsUpdates\UpdateEndpoints-Local.ps1 is missing locally. Please place that script file here and hit any key to re-check." -ForegroundColor Red
        }
    } WHILE ($doneCheckingLocal -ne $true)
    $doneCheckingLocal = $false
    DO {
        # Check that all required updates are downloaded
        $localUpdates = Get-ChildItem -Path C:\WindowsUpdates\* -Include "*.msu", "*.exe"; # Get local downloaded updates
        $missingUpdatesMasterList | % { # loop through each unique missing update and check if they've been downloaded
            [bool]$match = $false;
            $location = "";
            foreach ($localUpdate IN $localUpdates) {
                $localUpdateKB = $localUpdate.Name.split(".")[0]
                If (($localUpdateKB[0] -eq "(") -and ($localUpdateKB[-1] -eq ")")) {
                    $localUpdateKB = $localUpdateKB.Replace("(", "");
                    $localUpdateKB = $localUpdateKB.Replace(")", "");
                }
                if ($localUpdateKB -eq $_.KB) { $match = $true; $location = $localUpdate.FullName}
            }
            if ($match -eq $true) {
                if ($location -match ".exe") {
                    write-host "$($_.Title) downloaded at $($location) - will copy to server but must be manually installed" -ForegroundColor Yellow;
                } else {
                    write-host "$($_.Title) downloaded at $($location)" -ForegroundColor Green;
                }
                
            }
            else {
                write-host "$($_.Title) missing" -ForegroundColor Red;
            }
        }
        $confirmation = read-host "Note the items missing. Recheck local downloads (y) or move on to copying these to remote endpoints (n)?"
        if ($confirmation -eq 'y') {
            Write-Host "Scanning local folder for updates again...";
            continue;

        } elseif ($confirmation -eq 'n') {
            $doneCheckingLocal = $true;
        }

    } WHILE ($doneCheckingLocal -ne $true)
    Write-Host "Moving on to copying updates to remote endpoints...";

    # Loop through each server, check if local updates exist, and copy any missing updates
    ForEach ($server IN $serverMasterList)
    {
        $test = Test-Path -path "\\$($server.ServerName)\c$\WindowsUpdates";
        If ($test -eq $true) {} else {
            Write-Host "Remote folder does not exist; creating..."
            New-Item -ItemType directory -Path "\\$($server.ServerName)\c$\WindowsUpdates"
        }
        # Check if remote machine has the local ps1 script; prompt for overwrite if exists
        $remoteScript = Test-Path "\\$($server.ServerName)\c$\WindowsUpdates\UpdateEndpoints-Local.ps1"
        If ($remoteScript -eq $true) {
            $response = Read-Host "Remote script exists; overwrite (y/n)?";
            if ($response -eq "y") {
                Write-Host "Overwriting remote script with new version..."
                Copy-Item -Path "C:\WindowsUpdates\UpdateEndpoints-Local.ps1" -Destination "\\$($server.ServerName)\c$\WindowsUpdates\UpdateEndpoints-Local.ps1" -Force
            } else {
                write-host "Leaving remote script as-is";
            }
        } else {
            Write-Host "Remote script does not exist; copying..."
            Copy-Item -Path "C:\WindowsUpdates\UpdateEndpoints-Local.ps1" -Destination "\\$($server.ServerName)\c$\WindowsUpdates\UpdateEndpoints-Local.ps1" -Force
        }
        $remoteLocalUpdates = Get-ChildItem "\\$($server.ServerName)\c$\WindowsUpdates\*" -Include "*.msu", "*.exe";
        $server.MissingKB | % { # loop through each unique missing update and check if they've been copied to server

            [bool]$match = $false;
            $location = "";
            foreach ($remoteUpdate IN $remoteLocalUpdates) {
                $remoteUpdateKB = $remoteUpdate.Name.split(".")[0]
                If (($remoteUpdateKB[0] -eq "(") -and ($remoteUpdateKB[-1] -eq ")")) {
                    $remoteUpdateKB = $remoteUpdateKB.Replace("(", "");
                    $remoteUpdateKB = $remoteUpdateKB.Replace(")", "");
                }
                if ($remoteUpdateKB -eq $_) {$match = $true; $location = $remoteUpdate.FullName}
            }
            if ($match -eq $true) { write-host "$($_) present on $($server.ServerName)" -ForegroundColor Green; }
            else {
                $copyFile = Get-ChildItem -Path "C:\WindowsUpdates" -recurse -Filter "*$($_)*"
                if (!([string]::IsNullOrEmpty($copyFile.FullName))) { # This is a messy way to do it, but will require a refactor to change. Current;y basing the loop on the "MissingKB" property of the object instead of basing this on the local updates
                    write-host "Copying $($_) to $($server.ServerName)..."
                    write-host "$($_) missing on $($server.ServerName)" -ForegroundColor Red;
                    write-host "Copying $($copyFile.FullName) to \\$($server.ServerName)\c$\WindowsUpdates\$($copyFile)"
                    cmd /c copy /z $copyFile.FullName "\\$($server.ServerName)\c$\WindowsUpdates\$($copyFile)" # using cmd so I can see the progress of the file
                }
            }
        }
    }

    write-host "All updates have been copied to the remote endpoints. `n"
    $response = read-host "To install patches on endpoints and reboot them (you will be prompted - not automatic), press any key. Otherwise, exit this window."
    $serverInvokeList = @();
    $serverMasterList | % {
        if ($_.MissingKB.count -gt 0) {
            $serverInvokeList += $_.ServerName;
        }
    }
    foreach ($server IN $serverInvokeList) {
        Start-Process -FilePath 'PowerShell.exe' -ArgumentList '-NoExit',"-command `"Write-Host 'C:\WindowsUpdates\UpdateEndpoints-Local.ps1'; Enter-PSSession -ComputerName $server`";"
    }

    # Start testing for all endpoints reboot cycle
    $endpointRebootCycle = @();
    foreach ($server IN $serverInvokeList) {
        $thisServer = [pscustomobject]@{
            ServerName = $server;
            Status = "UP";
        };
        $endpointRebootCycle += $thisServer;
    }
    $allEndpointsRebooted -ne $false
    DO {
        foreach ($endpoint IN $endpointRebootCycle) {
            $test = Test-Connection -ComputerName $endpoint.ServerName -quiet;
            if ($test) {
                # Status of endpoint is up
                $status = "UP";
                $foregroundColor = "Green";
            } else {
                # Status of endpoint is down
                $status = "DOWN";
                $foregroundColor = "Red";
            }
            if ($status -eq $endpoint.Status) {
                # No change, don't change or print anything
            }
            else {
                # Change detected, write to console
		        $endpoint.Status = $status;
                Write-Host "$($endpoint.ServerName) status changed to $($endpoint.Status)" -ForegroundColor $foregroundColor;
            }
        }
    } WHILE ( $allEndpointsRebooted -ne $true )
}
