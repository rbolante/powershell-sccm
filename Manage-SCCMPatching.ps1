###############################################################################
#  Manage-SCCMPatching
#  Version-ID: 1.00
#  Author: Ronald Bolante
#  Copyright (c) 2016. All rights reserved.
###############################################################################
Param (
    [Parameter(Mandatory = $False)][ValidateSet(
        'PVT',
        'Patch',
        'Reboot',
        'Shutdown',
        'RunCycle',
        'GetStatus',
        'GetPending',
        'CleanupCache',
        'PreServiceSnap',
        'PostServiceSnap',
        'GetPatchByID',
        'GetPatchByTitle',
        'GetPatchFromDate',
        'GenerateServerLists',
        'RefreshComplianceState')][String]$Operation = 'GetPending',
    [Parameter(Mandatory = $False)][ValidateSet(
        'All',
        'HardwareInventory',
        'SoftwareInventory',
        'FileCollection',
        'DiscoveryDataCollection',
        'MachinePolicyRetrievalAndEvaluation',
        'SofwareMeteringUsageReport',
        'WindowsInstallerSourceListUpdate',
        'SoftwareUpdatesDeployment',
        'SoftwareUpdatesScan',
        'ApplicationDeploymentEvaluation')][String]$CycleName,
    [Parameter(Mandatory = $False)][String[]]$ComputerName,
    [Parameter(Mandatory = $False)][String[]]$ClusterName,
    [Parameter(Mandatory = $False)][String]$DeviceListFilePath,
    [Parameter(Mandatory = $False)][String]$ArticleID,
    [Parameter(Mandatory = $False)][String]$Filter,
    [Parameter(Mandatory = $False,HelpMessage="Format: YYYY-MM-DD")][String]$SnapDate = $(Get-Date).Date.ToString('yyyy-MM-dd'),
    [Parameter(Mandatory = $False)][Switch]$SendToEmail,
    [Parameter(Mandatory = $False)][Switch]$Force,
    [Parameter(Mandatory = $False)][Switch]$Whatif
)

$ErrorActionPreference = "SilentlyContinue"

If (-Not [Environment]::UserInteractive) {
    $VerbosePreference = "SilentlyContinue"
}

$Global:ParentDir = Split-Path $SCRIPT:MyInvocation.MyCommand.Path -Parent
$Global:ScriptName = (Split-Path $SCRIPT:MyInvocation.MyCommand.Path -Leaf).Split('.')[0]
$Global:LogFilePath = $($Global:ParentDir + '\' + $Global:ScriptName + '_' + $Operation + '_' + $(Get-Date -UFormat "%Y-%m-%d_%H-%M") + '.log')

$PSEmailServer = '<server-name>'
$EmailSender = "Patch Manager <PatchManager@<domain-name>"

$EvalStateHash = @{
    0 = 'Available'; # None
    1 = 'Available?';
    2 = 'Submitted';
    3 = 'Detecting';
    4 = 'PreDownload';
    5 = 'Downloading';
    6 = 'Waiting to install';
    7 = 'Installing';
    8 = 'Requires restart'; # Pending soft reboot
    9 = 'Pending hard reboot';
    10 = 'Waiting for reboot';
    11 = 'Pending verification';
    12 = 'Installation complete';
    13 = 'Error';
    14 = 'Waiting for service window';
    15 = 'Waiting for user logon';
    16 = 'User logoff';
    17 = 'Waiting for user logon';
    18 = 'Waiting user reconnect';
    19 = 'Pending user logoff';
    20 = 'Pending update';
    21 = 'Waiting for retry';
    22 = 'Waiting for Pres Mode Off';
    23 = 'Wait for orchestration';
}

$ClientCycleHash = @{
    HardwareInventory = "{00000000-0000-0000-0000-000000000001}";
    SoftwareInventory = "{00000000-0000-0000-0000-000000000002}";
    FileCollection = "{00000000-0000-0000-0000-000000000010}";
    DiscoveryDataCollection = "{00000000-0000-0000-0000-000000000003}";
    MachinePolicyRetrievalAndEvaluation = "{00000000-0000-0000-0000-000000000021}";
    SofwareMeteringUsageReport = "{00000000-0000-0000-0000-000000000022}";
    # UserPolicyRetrievalAndEvaluation = "{00000000-0000-0000-0000-000000000026}";
    WindowsInstallerSourceListUpdate = "{00000000-0000-0000-0000-000000000032}";
    SoftwareUpdatesDeployment = "{00000000-0000-0000-0000-000000000108}";
    SoftwareUpdatesScan	= "{00000000-0000-0000-0000-000000000113}";
    ApplicationDeploymentEvaluation = "{00000000-0000-0000-0000-000000000121}"
}


Function Write-Batch
{
    Param(
        [Parameter(Mandatory = $False)][ValidateSet('Error', 'Warning', 'Info')][String]$Level = 'Info',
        [Parameter(Mandatory = $False)][ConsoleColor]$ForegroundColor = [ConsoleColor]::White,
        [Parameter(Mandatory = $False)][Switch]$NoNewline,
        [Parameter(Mandatory = $True)][String]$Message = " "
    )
    $LogDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    If ([Environment]::UserInteractive) {
        Switch ($Level) {
            'Error' {
                If ($NoNewline.IsPresent) {
                    Write-Host -NoNewline -ForegroundColor Red "$Message"
                } Else {
                    Write-Host -ForegroundColor Red "$Message"
                }
            }
            'Warning' {
                Write-Warning -Message "$Message"
            }
            Default {
                If ($NoNewline.IsPresent) {
                    Write-Host -NoNewline -ForegroundColor $ForegroundColor "$Message"
                } Else {
                    Write-Host -ForegroundColor $ForegroundColor "$Message"
                }
            }
        }
        If ($Message.Trim().Length -gt 0) {
            $($LogDate + ' : ' + $Level.ToUpper() + ' : ' + $([Security.Principal.WindowsIdentity]::GetCurrent().Name) + ' : ' + $Message) | Out-File $Global:LogFilePath -Force -Append
        }
    } Else {
        If ($Message.Trim().Length -gt 0) {
            $($LogDate + ' : ' + $Level.ToUpper() + ' : ' + $([Security.Principal.WindowsIdentity]::GetCurrent().Name) + ' : ' + $Message) | Out-File $Global:LogFilePath -Force -Append
        }
    }
}


Function _PressAnyKey ()
{
    Write-Host -NoNewline "Press any key to continue..."
    $KeyPress = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}


Function _GetCMErrorMessage ([Int64]$ErrorCode)
{
    # TODO: It is bad to hard-code file paths - FIX THIS!!!
    [void][System.Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\SrsResources.dll")
    Return [SrsResources.Localization]::GetErrorMessage($ErrorCode, "en-US")
}


Function _GetPatchStatus ([String[]]$DeviceList, [String]$FilterString)
{
    $OutList = @()
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Processing $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {

                # Get pending updates status
                $KBList = Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate" -Namespace "ROOT\ccm\ClientSDK" -ComputerName $ComputerName
                $KBList | % {
                    $UpdateInfo = $_
                    $Percentage = '-'
                    $ErrorMessage = '-'
                    Switch ($UpdateInfo.EvaluationState) {
                        7 { $Percentage = $([String]$UpdateInfo.PercentComplete + '%') }
                        13 { $ErrorMessage = _GetCMErrorMessage -ErrorCode $UpdateInfo.ErrorCode }
                        Default { }
                    }
                    $Title = $UpdateInfo.Name
                    $TitleExtended = $UpdateInfo.Name
                    If ($UpdateInfo.Name.Length -gt 80) {
                        $Title = $($UpdateInfo.Name.ToString().SubString(0, 60) + " ...")
                    }

                    $OutList += New-Object -TypeName PSCustomObject -Property @{
                        ComputerName = $ComputerName
                        # ArticleID = $ArticleID
                        ArticleID = $UpdateInfo.ArticleID
                        Bulletin = $UpdateInfo.BulletinID
                        Title = $Title
                        TitleExtended = $TitleExtended
                        JobState = $EvalStateHash[[Int]$UpdateInfo.EvaluationState]
                        PercentComplete = $Percentage
                        ErrorMessage = $ErrorMessage
                    }
                }

                Write-Batch -ForegroundColor Green -Message "done."

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

        } Catch {
           Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    If ($FilterString.Length -gt 0) {
        $OutList | Where-Object {
            $_.TitleExtended -match $FilterString
        } | Sort-Object ComputerName | Select-Object ComputerName, Bulletin, ArticleID, Title, JobState, PercentComplete, ErrorMessage | FT -AutoSize
    } Else {
        $OutList | Sort-Object ComputerName | Select-Object ComputerName, Bulletin, ArticleID, Title, JobState, PercentComplete, ErrorMessage | FT -AutoSize
    }
}


Function _GetPendingUpdates ([String[]]$DeviceList)
{
    $OutList = @()
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Processing $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                # Get pending updates
                $KBList = (Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate" -Namespace "ROOT\ccm\ClientSDK" -ComputerName $ComputerName).ArticleID
                $OutList += New-Object -TypeName PSCustomObject -Property @{
                    ComputerName = $ComputerName
                    PendingUpdates = $KBList.Count
                    KBArticles = $KBList -join ", "
                }
                Write-Batch -ForegroundColor Green -Message "done."

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

        } Catch {
           Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    If (($OutList | Where-Object { $_.PendingUpdates -eq 0 }).Count -ne 0) {
        Write-Host
        Write-Batch -Message "* No pending updates on the following devices:"

        $OutList | Where-Object {
            $_.PendingUpdates -eq 0
        } | % {
            Write-Batch -Message $_.ComputerName
        }
    }

    Write-Host
    Write-Batch -Message "* Devices with pending updates:"
    $OutList | Select-Object ComputerName, PendingUpdates, @{Label='KBArticles'; Expression={ $_ | ForEach { $_.KBArticles }}} | Where-Object { 
        $_.PendingUpdates -ne 0
    }

    If ($SendToEmail.IsPresent) {
        Try {
            $HtmlHead = "<style>"
            $HtmlHead += "BODY{font-family: Arial; font-size: 10pt;}"
            $HtmlHead += "TABLE{border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}"
            $HtmlHead += "TH{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:#00ced1}"
            $HtmlHead += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:white}"
            $HtmlHead += "</style>"

            $OutList | Select-Object ComputerName, PendingUpdates, KBArticles | ConvertTo-Html -Head $HtmlHead | Out-File -FilePath "$Global:ParentDir\pending-updates.html" -Force
    
            $HtmlBody = Get-Content -Path "$Global:ParentDir\pending-updates.html" -Raw
            Send-MailMessage -From $EmailSender -To $RecipientList -Subject $("Pending Updates Report - " + $(Split-Path -Path $DeviceList -Leaf).Split('.')[0]) -Body $HtmlBody -BodyAsHtml

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }    
    }
}


Function _PatchDevices ([String[]]$DeviceList, [String]$KBNumber)
{
    $DeviceList | % {
        Try {
            $ComputerName = $_
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                [System.Management.ManagementObject[]]$MissingUpdates = @(Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate WHERE ComplianceState = 0 AND EvaluationState = 0" `
                    -Namespace "ROOT\CCM\ClientSDK" -ComputerName $ComputerName)
                # Write-host missing updates = $MissingUpdates.Count
                If ($MissingUpdates.Count -gt 0) {
                    If ($KBNumber.Length -gt 0) {
                        Write-Batch -NoNewline -Message "Triggering $($KBNumber) software update on $ComputerName... "
                    } Else {
                        Write-Batch -NoNewline -Message "Triggering $($MissingUpdates.Count) software updates on $ComputerName... "
                    }
                    $ResultObject = $Null
                    $ResultObject = Get-WmiObject -Class CCM_SoftwareUpdatesManager -Namespace "ROOT\CCM\ClientSDK" -List -ComputerName $ComputerName
                    If ($ResultObject -ne $Null) {
                        If ($KBNumber.Length -gt 0) {
                            $UpdateToApply = $MissingUpdates | Where-Object { $_.ArticleID -eq $KBNumber }
                            If ($UpdateToApply -ne $Null) {
                                $ResultObject.InstallUpdates($UpdateToApply) | Out-Null
                                Write-Batch -ForegroundColor Green -Message "done."
                            } Else {
                                Write-Batch -ForegroundColor Red -Message "no such update."
                            }
                        } Else {
                            $ResultObject.InstallUpdates($MissingUpdates) | Out-Null
                            Write-Batch -ForegroundColor Green -Message "done."
                        }
                    } Else {
                        Write-Batch -Level Error -Message "WMI error."
                    }
                }
            }
        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
}


Function _PatchDevicesSpecific ([String[]]$DeviceList, [String]$KBNumber)
{
    $DeviceList | % {
        Try {
            $ComputerName = $_
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                [System.Management.ManagementObject[]]$MissingUpdates = @(Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate WHERE ComplianceState = '0'" `
                    -Namespace "ROOT\CCM\ClientSDK" -ComputerName $ComputerName)
                If ($MissingUpdates.Count -gt 0) {
                    Write-Batch -NoNewline -Message "Triggering $($MissingUpdates.Count) software updates on $ComputerName... "
                    $ResultObject = $Null
                    $ResultObject = Get-WmiObject -Class CCM_SoftwareUpdatesManager -Namespace "ROOT\CCM\ClientSDK" -List -ComputerName $ComputerName
                    If ($ResultObject -ne $Null) {
                        $UpdateToApply = $MissingUpdates | Where-Object { $_.ArticleID -eq $KBNumber }
                        If ($UpdateToApply -ne $Null) {
                            $ResultObject.InstallUpdates($UpdateToApply) | Out-Null
                            Write-Batch -ForegroundColor Green -Message "done."
                        } Else {
                            Write-Batch -ForegroundColor Green -Message "No such update - $($KBNumber)"                            
                        }
                    } Else {
                        Write-Batch -Level Error -Message "WMI error."
                    }
                }
            }
        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
}


Function _RebootMachine ([String[]]$DeviceList, [Switch]$Force)
{   
    $Answer = _PromptUser -PromptString "Are you sure you want to continue? [YyNn]" -Choices "YyNn"
    If ($Answer.ToLower() -eq 'y') {
        $DeviceList | % {
            Try {
                $ComputerName = $_
                $KBList = Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate" -Namespace "ROOT\ccm\ClientSDK" -ComputerName $ComputerName
                $KBList | % {
                    $UpdateInfo = $_
                    $JobState = $EvalStateHash[[Int]$UpdateInfo.EvaluationState]
                    If ($JobState -eq 'Installing') { 
                        Write-Host
                        Write-Batch -NoNewline -Message "$($UpdateInfo.ArticleID) is installing... "
                        Write-Batch -Level Error -Message "Aborting restart."
                        Exit
                    }
                }
                Write-Batch -NoNewline -Message "Restarting $ComputerName... "
                If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                    If ($Force.IsPresent -Or ($(_CheckPendingReboot -ComputerName $ComputerName) -eq $True)) {
                        Restart-Computer -ComputerName $ComputerName -Force
                        Write-Batch -ForegroundColor Green -Message "done."
                    } Else {
                        Write-Batch -ForegroundColor DarkYellow -Message "not on pending reboot."
                    }
                } Else {
                    Write-Batch -Level Error -Message "ping failed."
                }
            } Catch {
                Write-Batch -Level Error -Message $_.Exception.Message
            }
        }
    }
}


Function _ShutdownMachine ([String[]]$DeviceList, [Switch]$Force)
{   
    $Answer = _PromptUser -PromptString "Are you sure you want to continue? [YyNn]" -Choices "YyNn"
    If ($Answer.ToLower() -eq 'y') {
        $DeviceList | % {
            Try {
                $ComputerName = $_
                $KBList = Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate" -Namespace "ROOT\ccm\ClientSDK" -ComputerName $ComputerName
                $KBList | % {
                    $UpdateInfo = $_
                    $JobState = $EvalStateHash[[Int]$UpdateInfo.EvaluationState]
                    If ($JobState -eq 'Installing') { 
                        Write-Host
                        Write-Batch -NoNewline -Message "$($UpdateInfo.ArticleID) is installing... "
                        Write-Batch -Level Error -Message "Aborting shutdown."
                        Exit
                    }
                }
                Write-Batch -NoNewline -Message "Shutting down $ComputerName... "
                If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                    If ($Force.IsPresent -Or ($(_CheckPendingReboot -ComputerName $ComputerName) -eq $True)) {
                        Stop-Computer -ComputerName $ComputerName -Force
                        Write-Batch -ForegroundColor Green -Message "done."
                    } Else {
                        Write-Batch -ForegroundColor DarkYellow -Message "not on pending reboot."
                    }
                } Else {
                    Write-Batch -Level Error -Message "ping failed."
                }
            } Catch {
                Write-Batch -Level Error -Message $_.Exception.Message
            }
        }
    }
}


Function _PreCheckServices ([String[]]$DeviceList)
{
    $DestPath = "$Global:ParentDir\Snapshots\$($(Get-Date).Date.ToString('yyyy-MM-dd'))"
    If (-Not (Test-Path -Path $DestPath)) {
        New-Item -Path $DestPath -ItemType Directory -Force | Out-Null
    }

    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Processing $ComputerName... "
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                Get-Service -ComputerName $ComputerName | Sort-Object ServiceName | Select-Object ServiceName, CanPauseAndContinue, `
                    CanShutdown, CanStop, Status | Export-Clixml -Path $($DestPath + '\' + $ComputerName + '-PreSnap.xml') -Force
                Write-Batch -ForegroundColor Green -Message "done."
            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    Write-Host
    Write-Batch -Message "Snapshots have been saved in '$($DestPath)'."
}


Function _PostCheckServices ([String[]]$DeviceList, [String]$DateToCheck)
{
    $DestPath = "$Global:ParentDir\Snapshots\$DateToCheck"
    If (-Not (Test-Path -Path $DestPath)) {
        Write-Batch -Level Error -Message "No such folder - $($DestPath)"
        Return
    }

    $DiffResults = @()

    $DeviceList | % {
        Try {
            $ComputerName = $_

            Write-Batch -NoNewline -Message "Processing $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                If (Test-Path -Path $($DestPath + '\' + $ComputerName + '-PreSnap.xml')) {
                    $PreSnapData = Import-Clixml -Path $($DestPath + '\' + $ComputerName + '-PreSnap.xml')

                    Get-Service -ComputerName $ComputerName | Sort-Object ServiceName | Select-Object ServiceName, CanPauseAndContinue, `
                        CanShutdown, CanStop, Status | Export-Clixml -Path $($DestPath + '\' + $ComputerName + '-PostSnap.xml') -Force

                    $PostSnapData = Import-Clixml -Path $($DestPath + '\' + $ComputerName + '-PostSnap.xml')

                    $PostSnapData | % {
                        $PostObject = $_
                        $PreObject = $PreSnapData | Where-Object { $_.ServiceName -eq $PostObject.ServiceName }
    
                        If ($PreObject -ne $Null) {
                            $ChangeFlag = $False
                            $CanPauseAndContinue = ''
                            $CanShutdown = ''
                            $CanStop = ''
                            $Status = ''

                            If ($PreObject.CanPauseAndContinue -ne $PostObject.CanPauseAndContinue) {
                                $ChangeFlag = $True
                                $CanPauseAndContinue = $PostObject.CanPauseAndContinue.ToString()
                            }

                            If ($PreObject.CanShutdown -ne $PostObject.CanShutdown) {
                                $ChangeFlag = $True
                                $CanShutdown = $PostObject.CanShutdown.ToString()
                            }

                            If ($PreObject.CanStop -ne $PostObject.CanStop) {
                                $ChangeFlag = $True
                                $CanStop = $PostObject.CanStop.ToString()
                            }

                            If ($PreObject.Status -ne $PostObject.Status) {
                                $ChangeFlag = $True
                                $Status = $PostObject.Status.ToString()
                            }

                        } Else {
                            $ChangeFlag = $True
                            $Status = 'NewlyAdded'
                        }

                        If ($ChangeFlag -eq $True) {
                            $DiffResults += New-Object -TypeName PSCustomObject -Property @{
                                HostName = $ComputerName
                                ServiceName = $PostObject.ServiceName
                                CanPauseAndContinue = $CanPauseAndContinue
                                CanShutdown = $CanShutdown
                                CanStop = $CanStop
                                Status = $Status
                            }
                        }
                    }

                    Write-Batch -ForegroundColor Green -Message "done."

                } Else {
                    Write-Batch -ForegroundColor DarkYellow -Message "no pre-snapshot data available."
                }

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    If ($DiffResults.Count -gt 0) {
        $DiffResults | Select-Object HostName, ServiceName, CanPauseAndContinue, CanShutdown, CanStop, Status | FT -AutoSize

        $CsvFileOut = $($Global:ParentDir + '\Snap-Results_' + $DateToCheck + '.csv')

        Write-Batch -Message "Discrepancies found: Results saved in '$($CsvFileOut)'."
        $DiffResults | Select-Object HostName, ServiceName, CanPauseAndContinue, CanShutdown, CanStop, Status | Export-Csv -Path $CsvFileOut -NoTypeInformation -Force

    } Else {
        Write-Batch -Message "No difference detected."
    }
}


Function _RefreshComplianceState ([String[]]$DeviceList)
{
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Host -NoNewline "Refreshing compliance state on $ComputerName... "
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                (New-Object -ComObject Microsoft.CCM.UpdatesStore).RefreshServerComplianceState()
                New-EventLog -LogName Application -Source SyncStateScript -ErrorAction SilentlyContinue
                Write-EventLog -LogName Application -Source SyncStateScript -EventId 555 -EntryType Information -Message "Sync State ran successfully"
            }
            Write-Batch -ForegroundColor Green -Message "done."

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
}


Function _GetPVTResults ()
{
    Param (
        [Parameter(Mandatory = $True)][String[]]$DeviceList,
        [Parameter(Mandatory = $False)][Switch]$OutToCsv,
        [Parameter(Mandatory = $False)][Switch]$OutToHtml
    )

    $WorkerScript = '.\Invoke-ServerPVT.ps1'
    $ResultsFilePrefix = '.\pvt-results'
    $StartTime = Get-Date
    $StatusList = @()

    # Go through the server list
    Try {
        $UserCred = (Get-Credential -UserName $([Security.Principal.WindowsIdentity]::GetCurrent().Name) -Message "Input password")

        $DeviceList | % {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Running PVT on $ComputerName... "
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $JobStatus = Start-Job -Name $_ -FilePath $WorkerScript -Credential $UserCred -ArgumentList $ComputerName 
                Write-Batch -Message "JobID = $($JobStatus.Id)."
            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }
        }

        # Wait for all to complete
        While (Get-Job -State "Running") {
            Start-Sleep 1
        }

        # Get all job results
        Get-Job | % {
            $StatusList += Receive-Job -Name $_.Name -Keep
        }

        # Cleanup
        Remove-Job * -Force

    } Catch {
        Write-Host -ForegroundColor Red $_.Exception.Message

    } Finally {
        $EndTime = Get-Date

        $StatusList | Select-Object HostName, UpTime, LastPatchDate, FailedWinEvents, OSDiskStatus, PendingUpdates, RebootStatus, `
            RdpService | Format-Table -AutoSize

        Write-Host "Duration:" (New-TimeSpan -Start $StartTime -End $EndTime).TotalSeconds "second(s)"

        If ($SendToEmail.IsPresent) {
            $HtmlHead = "<style>"
            $HtmlHead += "BODY{font-family: Arial; font-size: 9pt;}"
            $HtmlHead += "TABLE{border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}"
            $HtmlHead += "TH{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:#00ced1}"
            $HtmlHead += "TD{border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:white}"
            $HtmlHead += "</style>"

            $ResultsHtmlFile = $($ResultsFilePrefix + '.html')
            If (Test-Path -Path $ResultsHtmlFile -PathType Leaf) {
                Remove-Item -Path $ResultsHtmlFile -Force
            }

            $StatusList | Select-Object HostName, TimeChecked, UpTime, LastPatchDate, FailedWinEvents, OSDiskStatus, PendingUpdates, `
                RebootStatus, RdpService | ConvertTo-Html -Head $HtmlHead | Out-File -FilePath $ResultsHtmlFile -Force

            $HtmlBody = Get-Content -Path $ResultsHtmlFile -Raw
            Send-MailMessage -From $EmailSender -To $RecipientList -Subject $("PVT Results - " + $(Split-Path -Path $DeviceList -Leaf).Split('.')[0]) -Body $HtmlBody -BodyAsHtml
        }

        If ($OutToCsv.IsPresent) {
            $CsvFileName = $($($ResultsFilePrefix + '_') + $($(Split-Path -Path $DeviceList -Leaf).Split('.')[0]) + '.csv')
            $StatusList | Select-Object HostName, TimeChecked, UpTime, LastPatchDate, FailedWinEvents, OSDiskStatus, PendingUpdates, `
                RebootStatus, RdpService | Export-Csv -Path $CsvFileName -Force -NoTypeInformation

            Send-MailMessage -From $EmailSender -To $RecipientList -Subject $("PVT Results - " + $(Split-Path -Path $DeviceList -Leaf).Split('.')[0]) -Attachments $CsvFileName
        }
    }
}


Function _CheckPendingReboot ([String]$ComputerName)
{
    $CbsStatus = $Null
    $ChangeFlag = $False

    # Run RemoteRegistry service if it is stopped
    If ((Get-Service -Name RemoteRegistry -ComputerName $ComputerName).Status -eq 'Stopped') {
        Set-Service -Name RemoteRegistry -ComputerName $ComputerName -Status Running
        $ChangeFlag = $True
    }

    $RegBase = $Null
    $RegBase = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,$ComputerName)
    If ($RegBase -eq $Null) {
        Return 'InternalError'
    }

    $WuStatus = $RegBase.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
    $OSBuild = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName -Property BuildNumber -ErrorAction SilentlyContinue).BuildNumber
    If ($OSBuild -ge 6001) {
        $CbsStatus = $RegBase.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')
    }
    $RegBase.Close()

    # Stop RemoteRegistry again if it was stopped prior to processing
    If ($ChangeFlag -eq $True) {
        Get-Service -Name RemoteRegistry -ComputerName $ComputerName | Stop-Service -Force
    }

    $CcmStatus = ([WmiClass]"\\$ComputerName\root\ccm\clientsdk:CCM_ClientUtilities").DetermineIfRebootPending().RebootPending
    If ($WuStatus -Or $CbsStatus -Or $CcmStatus) {
        Return $True
    }

    Return $False
}


Function _RunMachinePolicy ([String[]]$DeviceList, [String]$CycleName)
{
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Running $CycleName cycle on $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                # Trigger machine policy update
                # $SmsClient = [wmiclass]"\\$ComputerName\root\ccm:SMS_Client"
                
                $SmsClient = Get-WmiObject -Class SMS_Client -Namespace 'root\ccm' -ComputerName $ComputerName -List
                If ($SmsClient -ne $Null) {
                    If ($CycleName -eq 'All') {
                        Write-Host
                        $ClientCycleHash.GetEnumerator() | % {
                            Write-Batch -NoNewline -Message "  + Triggering '$($_.Key)' cycle... "
                            $SmsClient.TriggerSchedule($_.Value) | Out-Null
                            Write-Batch -ForegroundColor Green -Message "done."
                        }

                    } Else {
                        $SmsClient.TriggerSchedule($ClientCycleHash[$CycleName]) | Out-Null
                        Write-Batch -ForegroundColor Green -Message "done."
                    }

                } Else {
                    Write-Batch -Level Error -Message "internal error."
                }

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }
            
            Write-Host

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
    Write-Host
}


Function _ExtracArticleIDFromTitle ([String]$Title) {
    If ($Title.Split('(').Count -gt 1) {
        Return $Title.ToString().Split('(')[1].Split(')')[0]
    }
    Return '-'
}


Function _GetPatchByTitle ([String[]]$DeviceList, [String]$FilterString)
{
    $OutputList = @()
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Checking patches on $ComputerName... "
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $Session = [Activator]::CreateInstance([Type]::GetTypeFromProgID("Microsoft.Update.Session", $ComputerName))
                $Searcher = $Session.CreateUpdateSearcher()
                $HistoryCount = $Searcher.GetTotalHistoryCount()            
                $UpdateHistory = $Searcher.QueryHistory(0, $HistoryCount)
                $UpdateHistory | Sort-Object -Unique Date -Descending | % {
                    $ArticleID = _ExtracArticleIDFromTitle -Title $_.Title
                    $DateInstalled = $_.Date.ToString().Split()[0]
                    $ActualDate = $_.Date
                    $Title = $_.Title
                    $OutputList += New-Object PSObject -Property @{
                        ComputerName = $ComputerName
                        ArticleID = $ArticleID
                        Title = $Title
                        Date = $ActualDate
                        DateInstalled = $DateInstalled
                    }
                }

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

            Write-Batch -ForegroundColor Green -Message "done."

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    $OutputList | Where-Object {
        $_.Title -match $FilterString
    } | Sort-Object -Unique Date -Descending | Select-Object ComputerName, ArticleID, Title, DateInstalled `
        | Export-Csv -Path ".\$(Get-Date -UFormat "%Y-%m-%d_%H-%M").log" -NoTypeInformation -Force
}


Function _GetPatchByID ([String[]]$DeviceList, [String]$KBNumber)
{
    $OutputList = @()
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Checking $KBNumber on $ComputerName... "
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $Session = [Activator]::CreateInstance([Type]::GetTypeFromProgID("Microsoft.Update.Session", $ComputerName))
                $Searcher = $Session.CreateUpdateSearcher()
                $HistoryCount = $Searcher.GetTotalHistoryCount()            
                $UpdateHistory = $Searcher.QueryHistory(0, $HistoryCount)
                $UpdateHistory | Sort-Object -Unique Date -Descending | % {
                    $ArticleID = _ExtracArticleIDFromTitle -Title $_.Title
                    $DateInstalled = $_.Date.ToString()
                    $ActualDate = $_.Date
                    $Title = $_.Title
                    $OutputList += New-Object PSObject -Property @{
                        ComputerName = $ComputerName
                        ArticleID = $ArticleID
                        Title = $Title
                        Date = $ActualDate
                        DateInstalled = $DateInstalled
                    }
                }

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

            Write-Batch -ForegroundColor Green -Message "done."

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    $OutputList | Where-Object {
        $_.ArticleID -eq $('KB' + $KBNumber)
    } | Select-Object ComputerName, ArticleID, Title, DateInstalled | FT -AutoSize
}


Function _GetPatchFromDate ([String[]]$DeviceList, [String]$FromDate)
{
    $OutputList = @()
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Checking patches on $ComputerName... "
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $Session = [Activator]::CreateInstance([Type]::GetTypeFromProgID("Microsoft.Update.Session", $ComputerName))
                $Searcher = $Session.CreateUpdateSearcher()
                $HistoryCount = $Searcher.GetTotalHistoryCount()            
                $UpdateHistory = $Searcher.QueryHistory(0, $HistoryCount)
                $UpdateHistory | Sort-Object -Unique Date -Descending | % {
                    $ArticleID = _ExtracArticleIDFromTitle -Title $_.Title
                    $DateInstalled = $_.Date.ToLocalTime().ToString()
                    $ActualDate = $_.Date.ToLocalTime()
                    $Title = $_.Title
                    $OutputList += New-Object PSObject -Property @{
                        ComputerName = $ComputerName
                        ArticleID = $ArticleID
                        Title = $Title
                        Date = $ActualDate
                        DateInstalled = $DateInstalled
                    }
                }

                Write-Batch -ForegroundColor Green -Message "done."

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }

    $OutputList | Sort-Object ComputerName | Where-Object {
        $_.Date -ge [DateTime]$FromDate
    }  | Select-Object ComputerName, ArticleID, DateInstalled, Title

}


Function _UnPatchDevices ([String[]]$DeviceList, [String]$KBNumber, [Switch]$ForceRestart)
{
    $DeviceList | % {
        Try {
            $ComputerName = $_
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {

                $CmdString = "cmd /c wusa.exe /quiet /uninstall /KB:$($KBNumber) /norestart /log:C:\temp\$($KBNumber)_Uninstall.txt"
                If ($ForceRestart.IsPresent) {
                    $CmdString = "cmd /c wusa.exe /quiet /uninstall /KB:$($KBNumber) /forcerestart /log:C:\temp\$($KBNumber)_Uninstall.txt"
                }

                Invoke-WmiMethod -computername $Computername -class Win32_Process -Name Create -ArgumentList $CmdString
            }

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
}


Function _CleanupCache ([String[]]$DeviceList)
{
    $LastMonth = 3
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Host "$ComputerName :"
            Get-ChildItem -Path $('\\' + $ComputerName + '\C$\Windows\ccmcache') -Directory | Where-Object {
                $_.LastWriteTime -lt (Get-Date).AddMonths(-$LastMonth)
            } | Sort-Object LastWriteTime -Descending | % {
                Write-Batch -NoNewline -Message "  + Deleting $($_.FullName) ($($_.LastWriteTime.ToString('yyyy-MM-dd')))... "
                $_.FullName | Remove-Item -Recurse -Force
                Write-Batch -ForegroundColor Green -Message "done."
            }
        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
}


Function _PromptUser ([String]$PromptString, [String]$Choices)
{
    Do {
        $Answer = Read-Host -Prompt $PromptString
    } While (-Not ($Choices.ToCharArray() | ? { $Answer -in $_ }))
    Return $Answer
}



#################
#   S T A R T   #
#################

Try {
    $DeviceList = @()

    If ($ComputerName.Count -gt 0) {
        $DeviceList = $ComputerName
    } Else {
        If ($DeviceListFilePath.Length -gt 0) {
            Get-Content -Path $DeviceListFilePath -ErrorAction Stop | % {
                # Skip comments and blank lines
                If (-Not ($_.ToString().StartsWith('#')) -And ($_.ToString().Length -ne 0)) {
                    $DeviceList += $_.ToString().Trim()
                }
            }
        }
    }

    Switch ($Operation) {
        # Get pending updates
        'GetPending' {
            _GetPendingUpdates -DeviceList $DeviceList
        }

        # Get patch status information
        'GetStatus' {
            _GetPatchStatus -DeviceList $DeviceList -FilterString $Filter
        }

        # Get patches by title
        'GetPatchByTitle' {
            _GetPatchByTitle -DeviceList $DeviceList -FilterString $Filter
        }

        # Get patches by title
        'GetPatchByID' {
            _GetPatchByID -DeviceList $DeviceList -KBNumber $ArticleID
        }

        # Get patches from date
        'GetPatchFromDate' {
            _GetPatchFromDate -DeviceList $DeviceList -FromDate $SnapDate
        }

        # Initiate patching
        'Patch' {
            _PatchDevices -DeviceList $DeviceList -KBNumber $ArticleID
        }

        # Initiate patching
        'UnPatch' {
            _UnPatchDevices -DeviceList $DeviceList -KBNumber $ArticleID
        }

        # Refresh compliance state
        'RefreshComplianceState' {
            _RefreshComplianceState -DeviceList $DeviceList
        }

        # PVT
        'PVT' {
            _GetPVTResults -DeviceList $DeviceList
        }

        # Reboot
        'Reboot' {
            _RebootMachine -DeviceList $DeviceList -Force:$Force
        }

        # Shutdown
        'Shutdown' {
            _ShutdownMachine -DeviceList $DeviceList -Force:$Force
        }

        # Create a snapshot of services
        'PreServiceSnap' {
            _PreCheckServices -DeviceList $DeviceList
        }

        # Conduct post services check
        'PostServiceSnap' {
            _PostCheckServices -DeviceList $DeviceList -DateToCheck $SnapDate
        }

        # Run Machine policy retrieval
        'RunCycle' {
            _RunMachinePolicy -DeviceList $DeviceList -CycleName $CycleName
        }

        # Conduct CCM cache cleanup
        'CleanupCache' {
            _CleanupCache -DeviceList $DeviceList
        }

        Default {}
    }        

} Catch {
    Write-Batch -Level Error -Message $_.Exception.Message
}
