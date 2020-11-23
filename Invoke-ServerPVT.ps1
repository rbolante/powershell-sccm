###############################################################################
#  Invoke-ServerPVT
#  Version-ID: 1.00
#  Author: Ronald Bolante
#  Copyright (c) 2016. All rights reserved.
###############################################################################
Param (
    [Parameter(Mandatory = $True)][String]$ComputerName
)

$ErrorActionPreference = "SilentlyContinue"

$Global:FilterXML = $Null
$Global:Uptime = $Null
$Global:DnsRecordFile = ".\DnsRecordTypeA.txt"


Function _CheckPendingReboot ([String]$ComputerName)
{
    $RegBase = $Null
    $CbsStatus = $Null
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
    # $CcmStatus = ([WmiClass]"\\$ComputerName\root\ccm\clientsdk:CCM_ClientUtilities").DetermineIfRebootPending().RebootPending
    If ($WuStatus -Or $CbsStatus -Or $CcmStatus) {
        Return 'Pending'
    }
    Return 'NotPending'
}


Function _CheckWinUpdateStatus ([String]$ComputerName)
{
    Try {
        $Session = [Activator]::CreateInstance([Type]::GetTypeFromProgID("Microsoft.Update.Session", $ComputerName))
        $Searcher = $Session.CreateUpdateSearcher()
        $Limit = $Searcher.GetTotalHistoryCount()            
        $UpdateHistory = $Searcher.QueryHistory(0, $Limit) | Where-Object { 
            $_.Date.ToShortDateString() -eq $Global:FilterDate.ToShortDateString()
        } | ForEach-Object -Process {
            If ($_.ResultCode -ne 2) {
                Return $False
            }
        }
        Return $True
    } Catch {
        Return $False
    }
}


Function _CheckPendingUpdates ([String]$ComputerName)
{
    Try {
        Return (Get-WmiObject -Query "SELECT * FROM CCM_SoftwareUpdate WHERE ComplianceState = '0'" -Namespace "ROOT\CCM\ClientSDK" -ComputerName $ComputerName).Count.ToString()
    } Catch {
        Return 'InternalError'
    }
}


Function _CheckWinEventLogs ([String]$ComputerName)
{
    $OSVersion = $Null
    # $OSVersion = (Get-ADComputer -Identity $ComputerName -Properties OperatingSystem).OperatingSystem
    $OSVersion = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName).Version
    If ($OSVersion -eq $Null) {
        Return 'InternalError'
    }
    $EventList = $Null
    Switch -Wildcard ($OSVersion) {
        { ($_ -like "6*") -or `
          ($_ -like "10*") } {
            $EventList = Get-WinEvent -FilterXml $Global:FilterXML -ComputerName $ComputerName | Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message
            Break 
        }
        { ($_ -like "5*") } { 
            $FilterDate = [System.Management.ManagementDateTImeConverter]::ToDmtfDateTime($Global:Uptime)
            $EventList = Get-WmiObject -Class win32_ntlogevent -ComputerName $ComputerName -Filter "(logfile='System') AND (type='Error') AND (TimeWritten >= '$FilterDate')" `
                 | Select-Object @{LABEL='TimeGenerated';EXPRESSION={$_.ConverttoDateTime($_.TimeGenerated)}}, EventCode, Type, SourceName, Message
            Break 
        }     
        Default { Return 'UnsupportedOS' }
    }
    If ($EventList -ne $Null) {
        # $EventList  | Format-Table -AutoSize
        $EventList  | Format-Table -AutoSize | Out-File ".\WinEvents-$ComputerName.txt" -Force
        # Invoke-Expression "& H:\Scripts\PowerShell\PVT\PAPVT-CheckWindowsEventLogs.ps1 -ComputerName $ComputerName"
        Return 'Found'
    }
    Return 'None'
}


Function _CheckRDPService ([String]$ComputerName)
{
    Return (Get-Service -Name TermService -ComputerName $ComputerName).Status
}


Function _CheckDnsRecordA ([String]$ComputerName)
{
    $DnsRecordList = $Null
    $DnsRecordList = Get-Content -Path $Global:DnsRecordFile | Where-Object { $_.ToString().Length -ne 0 } | % { $_.ToString().TrimEnd().ToLower() }
    If ($DnsRecordList | Where-Object { $_ -eq $ComputerName.ToString().ToLower() }) {
        Return 'Present'
    } Else {
        Return 'Missing'
    }
}


Function _GetUpTimeInfo ([String]$ComputerName)
{
    $Global:Uptime = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName | `
        Select-Object @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.LastBootUpTime)}}).LastBootUpTime
    If ($Global:Uptime -eq $Null) {
        Return 'InternalError'
    }
    $FilterDate = ($Global:Uptime.ToUniversalTime()).ToString("yyyy-MM-ddThh:mm:ss.fffZ")
    $Global:FilterXML = "
<QueryList>
  <Query Id=`"0`" Path=`"System`">
    <Select Path=`"System`">*[System[(Level=1  or Level=2 or Level=3) and TimeCreated[@SystemTime&gt;='$FilterDate']]]</Select>
  </Query>
</QueryList>"
    Return $Global:Uptime.ToString()
}


Function _GetLastPatchDate ([String]$ComputerName)
{
    Try {
        $Session = [Activator]::CreateInstance([Type]::GetTypeFromProgID("Microsoft.Update.Session", $ComputerName))
        $Searcher = $Session.CreateUpdateSearcher()
        $Limit = $Searcher.GetTotalHistoryCount()
        $UpdateHistory = $Searcher.QueryHistory(0, $Limit)
        $DateLastPatched = $UpdateHistory | Sort-Object -Unique Date -Descending | Select-Object @{LABEL='InstalledOn';EXPRESSION={([DateTime]($_.Date)).ToLocalTime()}} -First 1
        Return $DateLastPatched.InstalledOn.ToString().Split()[0]
    } Catch {
        Return 'InternalError'
    }
}


Function _CheckDiskSpace ([String]$ComputerName)
{
    # Set threshold to 3GB
    $ThresholdGB = [Int32]10
    $OutString = 'Good'
    $OsDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter "DeviceID = 'C:'"
    $FreeGB = "{0:N1}" -f ($OsDisk.FreeSpace / 1GB)
    If ([Int]$FreeGB -lt $ThresholdGB) {
        $OutString = ($FreeGB + "GB")
    }
    Return $OutString
}


#################
#   S T A R T   #
#################

$StatusList = @()

Try {
    $UpTimeInfo = $Null
    $WinUpdate = $False
    $PendingUpdates = $False
    $WinEvents = 'None'
    $DnsRecord = 'Missing'
    $RdpService = 'Stopped'
    $PendingReboot = 'InternalError'
    $LastPatchDate = 'None'

    # Check: ICMP echo
    If (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) {
        $ChangeFlag = $False

        # Run RemoteRegistry service if it is stopped
        If ((Get-Service -Name RemoteRegistry -ComputerName $ComputerName).Status -eq 'Stopped') {
            Set-Service -Name RemoteRegistry -ComputerName $ComputerName -Status Running
            $ChangeFlag = $True
        }

        # CHECK: Uptime
        $UpTimeInfo = _GetUpTimeInfo -ComputerName $ComputerName
        
        # CHECK: Windows Updates Status
        # $WinUpdate = _CheckWinUpdateStatus -ComputerName $ComputerName
        
        # CHECK: Pending Updates
        $PendingUpdates = _CheckPendingUpdates -ComputerName $ComputerName
        
        # CHECK: Event Logs Status
        $WinEvents = _CheckWinEventLogs -ComputerName $ComputerName
        
        # CHECK: Pending reboot
        $PendingReboot = _CheckPendingReboot -ComputerName $ComputerName
        
        # CHECK: Date last patched
        $LastPatchDate = _GetLastPatchDate -ComputerName $ComputerName

        # CHECK: Disk space (threshold - 3GB)
        $DiskSpaceLeft = _CheckDiskSpace -ComputerName $ComputerName

        # CHECK: RDP Service
        $RdpService = _CheckRDPService -ComputerName $ComputerName
        
        # Stop RemoteRegistry again if it was stopped prior to processing
        If ($ChangeFlag -eq $True) {
            Get-Service -Name RemoteRegistry -ComputerName $ComputerName | Stop-Service -Force
        }

    } Else {
        $UpTimeInfo = 'PingFailed';
        $PendingUpdates = '-';
        $WinEvents = '-';
        $PendingReboot = '-';
        $LastPatchDate = '-';
        $DiskSpaceLeft = '-';
        $RdpService = '-';
    }

    $StatusList = New-Object -TypeName PSCustomObject -Property @{
        HostName = $ComputerName;
        TimeChecked = $(Get-Date -Format "d/M/yyyy HH:mm:ss");
        UpTime = $UpTimeInfo;
        PendingUpdates = $PendingUpdates;
        FailedWinEvents = $WinEvents;
        RebootStatus = $PendingReboot;
        LastPatchDate = $LastPatchDate;
        OSDiskStatus = $DiskSpaceLeft;
        RdpService = $RdpService;
    }
        
} Catch {
    Write-Host -ForegroundColor Red $_.Exception.Message

} Finally {
    $StatusList | Select-Object *
    #  $StatusList | Out-File -FilePath '.\Results.txt' -Append -Force
    # | Select-Object HostName, TimeChecked, UpTime, Ping, `
    #    WinUpdatesStatus, NoPendingUpdates, WinEventsStatus, NoPendingReboot, RdpService, DNSRecord
    
    # $StatusList | Select-Object HostName, TimeChecked, UpTime, Ping, `
    #    WinUpdatesStatus, NoPendingUpdates, WinEventsStatus, NoPendingReboot, RdpService, DNSRecord >'.\Results.txt'
}