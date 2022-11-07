# Name: Repair-VssWriters
# Function: Enumerate all VSS writers on the machine and for each that's got a problem restart the associated service to stablise the writer
# WARNING: Ensure this doesn't run during a backup, as bad things could happen. You have been warned!
# Author / maintainer: Alex Heylin - alex@gmal.co.uk
# Thanks to https://heresjaken.com/resolve-vss-errors-without-rebooting-your-server/ for the initial version of this. 

$ServiceArray = @{
'ASR Writer' = 'VSS';
'BITS Writer' = 'BITS';
'Certificate Authority' = 'EventSystem';
'COM+ REGDB Writer' = 'VSS';
'DFS Replication service writer' = 'DFSR';
'DHCP Jet Writer' = 'DHCPServer';
'FRS Writer' = 'NtFrs';
'FSRM writer' = 'srmsvc';
'IIS Config Writer' = 'AppHostSvc';
'IIS Metabase Writer' = 'IISADMIN';
'Microsoft Exchange Replica Writer' = 'MSExchangeRepl';
'Microsoft Exchange Writer' = 'MSExchangeIS';
'Microsoft Hyper-V VSS Writer' = 'vmms';
'MSMQ Writer (MSMQ)' = 'MSMQ';
'MS Search Service Writer' = 'EventSystem';
'NPS VSS Writer' = 'EventSystem';
'NTDS' = 'EventSystem';
'OSearch VSS Writer' = 'OSearch';
'OSearch14 VSS Writer' = 'OSearch14';
'OSearch15 VSS Writer' = 'OSearch15';
'Registry Writer' = 'VSS';
'Shadow Copy Optimization Writer' = 'VSS';
'Sharepoint Services Writer' = 'SPWriter';
'SPSearch VSS Writer' = 'SPSearch';
'SPSearch4 VSS Writer' = 'SPSearch4';
'SqlServerWriter' = 'SQLWriter';
'System Writer' = 'CryptSvc';
'TermServLicensing' = 'TermServLicensing';
'WDS VSS Writer' = 'WDSServer';
'WIDWriter' = 'WIDWriter';
'Windows Server Storage VSS Writer' = 'WseStorageSvc';
'WMI Writer' = 'Winmgmt';
}

$errors = 0;
$restarts = 0;
$VssWriters = vssadmin list writers | Select-String -Context 0,4 'writer name:' 

$VssWritersErrors = $VssWriters | ? {$_.Context.PostContext[3].Trim() -ne "Last error: No error"} 

If ($VssWritersErrors -ne $null) {
    $VssWritersErrors | Select Line | %{$_.Line.tostring().Split("'")[1]}| ForEach-Object {
        If ($ServiceArray.Item($_) -ne $null -and $ServiceArray.Item($_) -ne "") {
            try {
                Restart-Service $ServiceArray.Item($_) -Force ; 
                $restarts++ ;
                Write-Output "VSS Writer $($_) needed repair. Restarted $($ServiceArray.Item($_)) service." ;
            } catch {
                $errors++ ;
                Write-Warning "VSS Writer $($_) needed repair. Error occured restarting $($ServiceArray.Item($_)) service. $($_.Exception.Message)" ;
            }
        } else {
            Write-Warning "VSS Writer $($_) need repair. Unable to identify service from array in script." ;
        }
    }
    Write-Output "OK: Some VSS Writers reported errors and were restarted" ;
} else {
    if ( $($VssWriters.count) -eq 0) {
        Write-Output "ERROR: No VSS writers found. This is an OS fault. Reboot (CHKDSK /F C: recommended) and retry. If this happens again, check the Application event log for source VSS"
        exit 1
    } else {
        Write-Output "OK: All registered VSS writers reported: No error" ;
    }
}
