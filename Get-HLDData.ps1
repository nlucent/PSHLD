<#
.SYNOPSIS
   Extracts data from HLD template 5.2+
.DESCRIPTION
   Extracts data from HLD template 5.2+
   v5.5+ without modification, v5.0-v5.5 with slight modifications
.PARAMETER <paramName>
   -hld <HLD file to read>
   -help
   -debug
.EXAMPLE
   digest-hld.ps1 -hld "c:\users\blah\some hld file" -fmohost all -objectlist default
#>

param([string]$hld, [switch]$help = $false, [switch]$debug=$false)

# Hardcoded vals (in case needed)
$cmoColcount = 6
$fmoColCount = 11
$migration = $false   
$allServers = @()   # array of server objects for hld 
$debuglogs = @()


if ($help) {
    help
    exit 0
}

if (-Not (test-path $hld)) {
    "test-path noped"
    exit 1
}

$file = Get-ChildItem $hld

# Create excel com object
$xl = new-object -comobject excel.application
$xl.visible = $false
$xl.displayalerts = $false

# Open specified file
if ($debug) { "Opening HLD ..." }
$wb = $xl.Workbooks.Open($hld)

# get version of current HLD
$hldparams = $wb.worksheets | where {$_.name -eq "Parameters"}
$hldver = $hldparams.cells.item(2,1).value()


function ConsumeServerDefs {
    if ($debug) { write-host "Doin' work ..." }
    # Examples
    $sheetName = "Server Definitions"
    # Index variables
    $fmoObjects = @{}   # Objects available in HLD
    $newServer = @{}    # Temp hash for new server object
    $fmoColumns = @()
    $cmoColumns = @()
        # Keep track of which server we are on
    $scount = 0
    
    if ([string]::IsNullOrEmpty($hldver)) {
        "HLD incompatible"
        exit 1
    }
    
    if ($debug) { "Found HLD version $hldver ..." }
    if ($hldver -lt 5.5) {
        
        # Storage location hash
        $storage = @{"ID" = 0; "Multiplier"=1; "Type"=2; "Usage"=3; "Format"=4; "Size"=5; "DataToMove"=6; "BackupType"=7; "BackupLength"=8; "Retention"=9 }
        # Storage option location hash
        $storageOptions = @{"Type" = 0; "Options" = 5}
        $storageTotals = @{"Total" = 5; "MigrateData" = 6 }
        $dbOptions = @{"Name" = 0; "Size" = 5}
    }
    else {
        
        # Storage location hash
        $storage = @{"ID" = 0; "Multiplier"=2; "Type"=3; "Usage"=4; "Format"=5; "Size"=7; "DataToMove"=9;}
        # Storage option location hash
        $storageOptions = @{"Type" = 0; "Options" = 5}
        $fsBackupOptions = @{"Type" = 0; "Retention" = 4; "Exclusions" = 6; "Comments" = 0}
        $dbBackupOptions = @{"Type" = 0; "Retention" = 4; "Comments" = 0}
        $storageTotals = @{"Total" = 5; "MigrateData" = 6 }
    }

    if ($debug) { write-host "Variables have been varied ..." }
    $ws = $wb.worksheets | where {$_.name -eq $sheetName}
    
    if ($debug) { Write-Host "Calculating cell count ..."}
    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count
    
    # Create $fmoObjects hash based on column 1 values starting with _
    if ($debug) { write-host "Indexes are being indexified ..."}
    for ($row = 1; $row -lt $rowCount; $row++) { 
        $cellval = $ws.cells.item($row,1).value()
        if ($cellval -match '^_') {
            $fmoObjects[$cellval] = $row 
        }
    }

    for ($col = 1; $col -lt $colCount; $col++) { 
        $cellval = $ws.cells.item(1,$col).value()
        if ($cellval -eq "FMO" -or $cellval -eq "New Demand") {
            $fmoColumns += ,$col
        } 
        elseif ($cellval -eq "CMO") {
            $cmoColumns += ,$col
        }
    }
    
    if ($debug) { write-host "Spreading my sheets ..." }
    # Get _cpu_cores from server named "blah"
    # $ws.cells.item($fmoObjects["_cpu_cores"], $serverNames["blah"]).Value()
    #
    # get size of disk1 from server "blah"
    # $ws.cells.item($fmoObjects["_storage1", $servernames["blah"] + $storage["Size"]]).Value()
    #
    # Get storage type from server "blah" disk 1
    # $ws.cells.items($fmoObjects["_storage1_details", $servernames["blah"] + $storageOptions["Options']]).Value()

    
    # Create array of server objects containing data from all $fmoObjects
    foreach ($col in $fmoColumns) {  
        $scount++
        $newServer = @{}
        $newServer["_column"] = $col
                
        $ocount = 0
        foreach ($obj in $fmoObjects.keys) {
            $ocount++
            if (-not $debug) { Write-Progress -Activity "Consuming data" -status "Server $scount of $fmoColumns.count - $obj" -percentComplete ((($scount * $ocount) / ($fmoColumns.count * $fmoObjects.count)) * 100) }
            $debuginfo = @{}
            
            switch -regex ($obj) {
                '^_storage\d+$' {
                    # This is a storage object .. storage1, storage2 etc. We add the values for each $storage type to the fmohost column # to get the value
                    foreach ($val in $storage.keys) {
                    # Assign each value to $newserver hash keyed on object name
                        $tmpName = $obj + "_" + $val
                        $tmpRow = $fmoObjects[$obj]
                        $tmpCol = $col + $($storage[$val])
                        $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                        if ($debug) {
                            $debuginfo["Object"] = $tmpName
                            $debuginfo["Value"] = "{0}" -f $tmpVal
                            $debuginfo["Cell"] = "{0}" -f "$tmpRow,$tmpCol"
                            #$debuglogs += ,$debuginfo
                            #"$tmpName -" + $ws.cells.item($fmoObjects[$obj],$col + $storage[$val]).Value() + "- (" + $fmoObjects[$obj] + "," + ($col + $storage[$val]) + ")" 
                        }
                        $newServer[$tmpName] = $tmpVal
                        if ([string]::IsNullOrEmpty($newServer[$tmpName])) {
                            $newServer[$tmpName] = $null
                        }
                        if ($debug) {
                            $Script:debuglogs += ,$debuginfo
                        }
                
                    }
                }

                '^_storage_totals$' {
                    # Total storage and migrating data
                    foreach ($val in $storageTotals.keys) {
                        $tmpName = $obj + "_" + $val
                        $tmpRow = $fmoObjects[$obj]
                        $tmpCol = $col + $storageTotals[$val]
                        $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                        if ($debug) {
                            $debuginfo["Object"] = $tmpName
                            $debuginfo["Value"] = "{0}" -f $tmpVal
                            $debuginfo["Cell"] = "{0}" -f "$tmpRow,$tmpCol"
                        }
                        $newServer[$tmpName] = $tmpVal
                        if ([string]::IsNullOrEmpty($newServer[$tmpName])) {
                            $newServer[$tmpName] = $null
                        }
                        if ($debug) {
                            $Script:debuglogs += ,$debuginfo
                        }
                
                    }
                }
                '^_storage\d+_details$' {
                    # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                    foreach ($val in $storageOptions.keys) {
                        $tmpName = $obj + "_" + $val
                        $tmpRow = $fmoObjects[$obj]
                        $tmpCol = $col + $storageOptions[$val]
                        $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                        if ($debug) {
                            $debuginfo["Object"] = $tmpName
                            $debuginfo["Value"] = "{0}" -f $tmpVal
                            $debuginfo["Cell"] = "{0}" -f "$tmpRow,$tmpCol"
                        }
                        $newServer[$tmpName] = $tmpVal
                        if ([string]::IsNullOrEmpty($newServer[$tmpName])) {
                            $newServer[$tmpName] = $null
                        }
                        if ($debug) {
                            $Script:debuglogs += ,$debuginfo
                        }
                    
                    }
                }
                '^_fs_backups$' {
                    # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                    foreach ($val in $fsBackupOptions.keys) {
                        $tmpName = $obj + "_" + $val
                        $tmpRow = $fmoObjects[$obj]
                        $tmpCol = $col + $fsBackupOptions[$val]
                        $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                        if ($debug) {
                            $debuginfo["Object"] = $tmpName
                            $debuginfo["Value"] = "{0}" -f $tmpVal
                            $debuginfo["Cell"] = "{0}" -f "$tmpRow,$tmpCol"
                        }
                        $newServer[$tmpName] = $tmpVal
                        if ([string]::IsNullOrEmpty($newServer[$tmpName])) {
                            $newServer[$tmpName] = $null
                        }
                        if ($debug) {
                            $Script:debuglogs += ,$debuginfo
                        }
                    
                    }
                }
                '^_db_backups$' {
                    # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                    foreach ($val in $dbBackupOptions.keys) {
                        $tmpName = $obj + "_" + $val
                        $tmpRow = $fmoObjects[$obj]
                        $tmpCol = $col + $dbBackupOptions[$val]
                        $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                        if ($debug) {
                            $debuginfo["Object"] = $tmpName
                            $debuginfo["Value"] = "{0}" -f $tmpVal
                            $debuginfo["Cell"] = "{0}" -f "$tmpRow, $tmpCol"
                            #$debuglogs += ,$debuginfo
                            #"$tmpName -" + $ws.cells.item($fmoObjects[$obj], $col + $dbBackupOptions[$val]).Value() + "- (" + $fmoObjects[$obj] + "," + ($col + $dbBackupOptions[$val]) + ")" 
                        }
                        $newServer[$tmpName] = $tmpVal
                        if ([string]::IsNullOrEmpty($newServer[$tmpName])) {
                            $newServer[$tmpName] = $null
                        }
                        if ($debug) {
                            $Script:debuglogs += ,$debuginfo
                        }
                    }
                }
                '^_fs_backups_comment$' {
                    # not done yet

                }
                '^_db_backups_comment$' {

                    # not done yet

                }
                '^_db\d+_info$' {

                    # not done yet

                }
                '^_tsna_installed_software\d+$' {
                    
                    # not done yet

                }
                default {
                    $tmpVal = $ws.cells.item($fmoObjects[$obj], $col).Value()
                    # this is for all other object requests
                    if ($debug) {
                        $debuginfo["Object"] = $obj
                        $debuginfo["Value"] = "{0}" -f $tmpVal
                        $debuginfo["Cell"] = "{0}" -f $($fmoObjects[$obj]) + ",$col"
                        #$debuglogs += ,$debuginfo
                        #"$obj -" + $ws.cells.item($fmoObjects[$obj], $col).Value() + "- (" + $fmoObjects[$obj] + ",$col" + ")" 
                    }
                    $newServer[$obj] = $tmpVal
                    if ([string]::IsNullOrEmpty($newServer[$obj])) {
                        $newServer[$obj] = $null
                    }
                    if ($debug) {
                        $Script:debuglogs += ,$debuginfo
                    }
                }
                
            }

        }

        $server = New-Object -TypeName PSObject -Property $newServer
        $Script:allServers += ,$server
        if ($debug) {
            # print pretty shit
            $Script:debuglogs.ForEach({[PSCustomObject]$_}) | sort -property object | Format-Table -AutoSize
        }
        if ($debug) { write-host "Server $scount is dead to me ..."  }

    }
}

function help {
    "Not yet implemented"
}

ConsumeServerDefs


$summary = @{}


foreach($s in $allServers) {
    $summary.serverCount = $allServers.count
    if ($s._platform -eq "CTZ - VM") {
        $summary.totalVCpus += $s._cpu_cores
        $summary.totalVRam += $s._ram
        $summary.totalVDisk += $s._storage_totals_total
        $summary.totalVData += $s._storage_totals_migratedata
    }
    elseif ($s._platform -eq "PH - Physical") {
        $summary.totalPCpus += $s._cpu_cores
        $summary.totalPRam += $s._ram
        $summary.totalPDisk += $s._storage_totals_total
        $summary.totalPData += $s._storage_totals_migratedata
    }

}

$summary.totalDisk = $summary.totalVDisk + $summary.totalPDisk
$summary.totalData = $summary.totalVData + $summary.totalPData
$global:d = new-object -TypeName PSObject -Property $summary

write-host "HLD: $file.Name"
write-host "Total CMO Servers: $cmoColumns.count"
write-host "Total FMO Servers: $allServers.Count"

if ($allServers.count -lt 5) {
    $allServers | select _fmo_dc,_fmo_hostname,_platform,_cpu_cores,_ram,_storage_totals_total,_storage_totals_migratedata
}

$d | select serverCount, totalDisk, totalData, totalVCpus, totalVRam, totalVDisk, totalPCpus, totalPRam, totalPDisk

# release the WS, dont know if this is necessary or not
$ws = $null
# Shut'er down
$wb.close()
