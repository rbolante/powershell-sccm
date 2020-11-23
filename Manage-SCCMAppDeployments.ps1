###############################################################################
#  Manage-SCCMAppDeployments
#  Version-ID: 1.00
#  Author: Ronald Bolante
#  Copyright (c) 2016. All rights reserved.
###############################################################################
Param (
    [Parameter(Mandatory = $True)][ValidateSet(
        'RunCycle',
        'Install',
        'Uninstall',
        'GetStatus',
        'GetAvailable')][String]$Operation = 'GetAvailable',
    [Parameter(Mandatory = $False)][ValidateSet(
        'All',
        'HardwareInventory',
        'SoftwareInventory',
        'FileCollection',
        'DiscoveryDataCollection',
        'MachinePolicyRetrievalAndEvaluation',
        'SofwareMeteringUsageReport',
        'UserPolicyRetrievalAndEvaluation',
        'WindowsInstallerSourceListUpdate',
        'SoftwareUpdatesDeployment',
        'SoftwareUpdatesScan',
        'ApplicationDeploymentEvaluation')][String]$CycleName,
    [Parameter(Mandatory = $False)][String]$ApplicationName = 'All',
    [Parameter(Mandatory = $False)][String[]]$ComputerName,
    [Parameter(Mandatory = $False)][String]$DeviceListFilePath,
    [Parameter(Mandatory = $False,HelpMessage="Format: YYYY-MM-DD")][String]$SnapDate = $(Get-Date).Date.ToString('yyyy-MM-dd'),
    [Parameter(Mandatory = $False)][Switch]$SendToEmail,
    [Parameter(Mandatory = $False)][Switch]$Whatif
)

$ErrorActionPreference = "SilentlyContinue"

If (-Not [Environment]::UserInteractive) {
    $VerbosePreference = "SilentlyContinue"
}

$Global:ParentDir = Split-Path $SCRIPT:MyInvocation.MyCommand.Path -Parent
$Global:ScriptName = (Split-Path $SCRIPT:MyInvocation.MyCommand.Path -Leaf).Split('.')[0]
$Global:LogFilePath = "$($Global:ParentDir)\$($Global:ScriptName)_$(Get-Date -UFormat "%Y-%m-%d_%H-%M").log"

$PSEmailServer = '<server-name>'
$EmailSender = "Patch Manager <PatchManager@<domain-name>"

$EvalStateHash = @{
    0 = 'No state information is available';
    1 = 'Application is enforced to desired/resolved state';
    2 = 'Application is not required on the client';
    3 = 'Application is available for enforcement. Content may/may not have been downloaded';
    4 = 'Application last failed to enforce (install/uninstall)';
    5 = 'Application is currently waiting for content download to complete';
    6 = 'Application is currently waiting for content download to complete';
    7 = 'Application is currently waiting for its dependencies to download';
    8 = 'Application is currently waiting for a service (maintenance) window';
    9 = 'Application is currently waiting for a previously pending reboot';
    10 = 'Application is currently waiting for serialized enforcement';
    11 = 'Application is currently enforcing dependencies';
    12 = 'Application is currently enforcing';
    13 = 'Application install/uninstall enforced and soft reboot is pending';
    14 = 'Application installed/uninstalled and hard reboot is pending';
    15 = 'Update is available but pending installation';
    16 = 'Application failed to evaluate';
    17 = 'Application is currently waiting for an active user session to enforce';
    18 = 'Application is currently waiting for all users to logoff';
    19 = 'Application is currently waiting for a user logon';
    20 = 'Application in progress, waiting for retry';
    21 = 'Application is waiting for presentation mode to be switched off';
    22 = 'Application is pre-downloading content (downloading outside of install job)';
    23 = 'Application is pre-downloading dependent content (downloading outside of install job)';
    24 = 'Application download failed (downloading during install job)';
    25 = 'Application pre-downloading failed (downloading outside of install job)';
    26 = 'Download success (downloading during install job)';
    27 = 'Post-enforce evaluation';
    28 = 'Waiting for network connectivity';
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
        [Parameter(Mandatory = $True)][String]$Message
    )
    If ($Message.Length -eq 0) {
        Return
    }
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
    [void][System.Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\SrsResources.dll")
    Return [SrsResources.Localization]::GetErrorMessage($ErrorCode, "en-US")
}


Function _GetAvailableApps ([String[]]$Devices)
{
    $OutList = @()
    $Devices | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Processing $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $WmiResult = Get-WmiObject -Class CCM_Application -Namespace 'root\CCM\ClientSDK' -ComputerName $ComputerName -Filter "ApplicabilityState = 'Applicable'"

                If ($WmiResult.Length -gt 0) {
                    $WmiResult | Where-Object {
                        $_.ResolvedState -eq 'Available' -and `
                        $_.InstallState -eq 'NotInstalled' 
                    } | % {
                        $OutList += New-Object -TypeName PSCustomObject -Property @{
                            ComputerName = $ComputerName
                            ApplicationName = $_.Name
                        }
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
    $OutList | Select-Object ComputerName, ApplicationName
}


Function _GetAppsStatus ([String[]]$Devices, [String]$AppName)
{
    $OutList = @()
    $Devices | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Processing $ComputerName... "
            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $WmiResult = Get-WmiObject -Class CCM_Application -Namespace 'root\CCM\ClientSDK' -ComputerName $ComputerName -Filter "ApplicabilityState = 'Applicable'"
                If ($WmiResult.Length -gt 0) {
                    $PercentComplete = $WmiResult.PercentComplete
                    If ($AppName -ne 'All') {
                        $WmiResult | Where-Object { $_.Name -like "$AppName*" } | % {
                            $OutList += New-Object -TypeName PSCustomObject -Property @{
                                ComputerName = $ComputerName
                                ApplicationName = $_.Name
                                InstallState = $_.InstallState
                                EvaluationState = $EvalStateHash[[Int]$_.EvaluationState]
                                ErrorMessage = _GetCMErrorMessage -ErrorCode $_.ErrorCode
                                PercentComplete = $PercentComplete
                            }
                        }
                    } Else {
                        $WmiResult | % {
                            $OutList += New-Object -TypeName PSCustomObject -Property @{
                                ComputerName = $ComputerName
                                ApplicationName = $_.Name
                                InstallState = $_.InstallState
                                EvaluationState = $EvalStateHash[[Int]$_.EvaluationState]
                                ErrorMessage = _GetCMErrorMessage -ErrorCode $_.ErrorCode
                                PercentComplete = $PercentComplete
                            }
                        }
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
    $OutList | Select-Object ComputerName, ApplicationName, InstallState, EvaluationState, ErrorMessage | FT -AutoSize
}


Function _DeployApps ([String[]]$Devices, [String]$AppName)
{
    $Devices | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Triggering install of $AppName on $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                $AppObject = Get-WmiObject -Class CCM_Application -Namespace 'root\CCM\ClientSDK' -ComputerName $ComputerName -Filter "ApplicabilityState = 'Applicable'" `
                    | Where-Object { $_.Name -like $AppName }
                If ($AppObject -ne $Null) {
                    $SmsClient = [wmiclass]"\\$ComputerName\root\CCM\ClientSDK:CCM_Application"
                    $SmsClient.Install($AppObject.Id, $AppObject.Revision, $AppObject.IsMachineTarget, 0, 'High', $False) | Out-Null
                    Write-Batch -ForegroundColor Green -Message "done."

                } Else {
                    Write-Batch -Level Error -Message "none triggered."
                }

            } Else {
                Write-Batch -Level Error -Message "ping failed."
            }

        } Catch {
            Write-Batch -Level Error -Message $_.Exception.Message
        }
    }
}


Function _RunMachinePolicy ([String[]]$DeviceList, [String]$CycleName)
{
    $DeviceList | % {
        Try {
            $ComputerName = $_
            Write-Batch -NoNewline -Message "Running $CycleName cycle on $ComputerName... "

            If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
                # Trigger machine policy update                
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




#################
#   S T A R T   #
#################

Try {

    $DeviceList = @()

    If ($ComputerName.Count -gt 0) {
        $DeviceList = $ComputerName
    } Else {
        Get-Content -Path $DeviceListFilePath -ErrorAction Stop | % {
            # Skip comments and blank lines
            If (-Not ($_.ToString().StartsWith('#')) -And ($_.ToString().Length -ne 0)) {
                $DeviceList += $_.ToString().Trim()
            }
        }
    }

    Switch ($Operation) {
        'GetAvailable' {
            # Get pending updates
            _GetAvailableApps -Devices $DeviceList
        }

        'GetStatus' {
            # Conduct post services check
            _GetAppsStatus -Devices $DeviceList -AppName $ApplicationName
        }

        { $_ -in 'Install', 'Uninstall' } {
            # Install/Uninstall
            _DeployApps -Devices $DeviceList -AppName $ApplicationName
        }

        'Reboot' {
            # Reboot
            _RebootMachine -Devices $DeviceList
        }

        'RunCycle' {
            # Run Machine policy retrieval
            _RunMachinePolicy -DeviceList $DeviceList -CycleName $CycleName
        }

        Default {}
    }        

} Catch {
    Write-Batch -Level Error -Message $_.Exception.Message
}

