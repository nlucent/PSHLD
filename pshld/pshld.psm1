<#
pshld.psm1 - Nick.Lucent at <redacted> v1
.SYNOPSIS
   Import data from HLD template v5.2+
.DESCRIPTION
   Creates PSCustomObject containing all data from the specified HLD file

.PARAMETER
    

.EXAMPLE
    $hld = import-hld c:\path\to\hld\template
    Import all available data from a single HLD

    $hld = import-hld c:\path\to\hld\template -force
    Import data from HLD, forcing refresh of any cached xml data

    $hld = import-hld c:\path\to\hld\template -force -cmoOnly
    ls *.xlsm | import-hld -force -cmoOnly | export-hld -cmoOnly
      
#>



function global:Import-HLD  {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)][Alias('FullName')][string[]]$hldpath,
        [parameter(Mandatory=$false, ValueFromPipeline=$false)][switch]$force,
        [parameter(Mandatory=$false, ValueFromPipeline=$false)][switch]$cmoOnly,
        [parameter(Mandatory=$false, ValueFromPipeline=$false)][switch]$fmoOnly
        )

    BEGIN {
        # Create excel com object
        $xl = new-object -comobject excel.application
        $xl.visible = $false
        $xl.displayalerts = $false
    }

    PROCESS {
        # Get first object from pipeline
        foreach($hld in $hldPath) {
            $file = Get-ChildItem $hld
            $allServers = @()   # array of server objects for hld 
            $thisHLD = new-object -type PSCustomObject 
            $newhash = (get-filehash $hld).hash

            # import cached HLD data
            if (test-path "$hld.hld") {
                $tempHLD = import-clixml "$hld.hld"
            }

            # reload if -force or if file hash doesnt match MD5 property of object
            if (($tempHLD.MD5 -eq $newhash) -and (-not $force) ) {
                $thisHld = $tempHLD
            }
            else {

                $thisHLD | add-member -type NoteProperty -name Name -value $file.Name -force

                # Make sure HLD file exists for import
                if (-not (test-path $hld)) {
                    "test-path noped"
                }

                # Open specified file
                write-debug "Opening HLD ..." 
                $wb = $xl.Workbooks.Open($hld)

                # get version of current HLD
                $hldparams = $wb.worksheets | where {$_.name -eq "Parameters"}
                $hldver = $hldparams.cells.item(2,1).value()

                # Read overview tab
               $thisHld = Import-Overview

                # read app requirements
                $thisHld = Import-AppReqs
                # Read server definitions
                $thisHld = Import-ServerDefs

                # release the worksheet, dont know if this is necessary or not
                $ws = $null
                
                # Shut'er down
                $wb.close()

                # Add hash to hld
                $thisHld |add-member -type NoteProperty -name MD5 -value $newhash -force

                # save xml version of hld
                $thisHld | export-clixml "$hld.hld"
            }
            # Return $thisHLD to caller
            $thisHLD
        }
    }

    END { 
    $xl.quit()
    }
}

#Import server definitions sheet from HLD
function Import-ServerDefs {
    $sheetName = "Server Definitions"
    # Index variables
    $hldObjects = @{}   # Objects available in HLD
    $newServer = @{}    # Temp hash for new server object
    $fmoColumns = @()   # Array of FMO server columns
    $cmoColumns = @()   # Array of CMO server columns
    $cmoServers = @()   # Array of all CMO servers in HLD
    $fmoServers = @()   # Array of all FMO servers in HLD
    
    #import HLD options
    $HLDOptions = import-clixml "$PSScriptRoot\options.hld"

    if ([string]::IsNullOrEmpty($hldver)) {
        write-error "HLD incompatible"
        break
    }
    
    write-debug "Found HLD version $hldver ..." 
    if ($hldver -lt 5.5) {
        $version = "v5"
    }
    else {
        $version = "v55"
    }

    write-debug "Variables have been varied ..." 
    $ws = $wb.worksheets | where {$_.name -eq $sheetName}
    
    write-debug "Calculating cell count ..."
    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count
    
    # Create $hldObjects hash based on column 1 values starting with _
    write-debug "Indexes are being indexified ..."
    for ($row = 1; $row -lt $rowCount; $row++) { 
        $cellval = $ws.cells.item($row,1).value()
        if ($cellval -match '^_') {
            $hldObjects[$cellval] = $row 
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
    
    write-debug "Spreading some sheets ..."

    
   # Keep track of which server we are on
   if (-not $fmoOnly) {
        $scount = 0
        foreach ($col in $cmoColumns) {  
            $scount++
            $cmo = Import-CMO $col $scount
            $cmoServers += ,$cmo
        } 
        $thisHLD | add-member -type NoteProperty -name CMO -value $cmoServers -force
    }

    if (-not $cmoOnly) {
        $scount = 0
        foreach ($col in $fmoColumns) {  
            $scount++
            $fmo = Import-FMO $col $scount
            $fmoServers += ,$fmo
        } 
        $thisHLD | add-member -type NoteProperty -name FMO -value $fmoServers -force
    }
    $ws = $null
    $thisHLD
}

# Import FMO data from server definitions sheet in HLD
function Import-FMO ($fCol, $scount) {
    # import FMO data to $hld.FMO()
    $thisServer = new-object PSObject
    $thisServer | add-member -type NoteProperty -name _column -value $col | add-member -type NoteProperty -name _disposition -value "FMO"
    $thisServer | add-member -type noteproperty -name _hld_name -value $thishld.name
    
    $newSoftware = @()
    $ocount = 0
    $col = $fCol

    # Create array of empty hashes for storage info
    $serverstorage = 0..15
    for ($i = 0; $i -lt 15; $i++) {
        $serverstorage[$i] = @{}
    }

    foreach ($obj in $hldObjects.keys | sort) {
        $ocount++
        if (-not $debug) { Write-Progress -Activity "Importing FMO data" -status "Server $scount of $($fmoColumns.count) - $obj" -percentComplete ((($scount * $ocount) / ($fmoColumns.count * $hldObjects.count)) * 100) }      

        switch -regex ($obj) {
            '^_storage\d+$' {
                # Retrieve the number from the object 
                $blah = $obj -match '\D+(\d+)\D*'
                $objNum = $matches[1]
                
                # Iterate over all storage keys and add to storage hash
                foreach ($val in $hldOptions.$version.storage.keys) {
                # Assign each value to $serverStorage
                    $tmpName = $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.$version.storage[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")){
                        $serverStorage[$objNum][$tmpName] = $tmpVal
                    }
                }
                foreach ($val in $hldOptions.$version.storageOptions.keys) {
                    $tmpName = $val
                    $tmpRow = $hldObjects[$obj] + 1
                    $tmpCol = $col + $hldOptions.$version.storageOptions[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if (-not [string]::IsNullOrEmpty($tmpVal) -and ($tmpval -ne "(Choose)") -and ($tmpVal -ne "(Select Storage Options)")) {
                        $serverStorage[$objNum][$tmpName] = $tmpVal
                    }
                }
                break
            }
            '^_storage\d+_details$' {
                break
            }
            '^_storage_totals$' {
                # Total storage and migrating data
                foreach ($val in $hldOptions.$version.storageTotals.keys) {
                    $tmpName = $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.$version.storageTotals[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")) {
                        #$thisServer | add-member -type NoteProperty -name $tmpName -value $tmpVal 
                        $serverStorage[0][$tmpName] = $tmpVal
                    }
                }
                break
            }
            '^_fs_backups$' {
                # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                foreach ($val in $hldOptions.$version.fsBackupOptions.keys) {
                    $tmpName = $obj + "_" + $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.$version.fsBackupOptions[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")) {
                        $thisServer | add-member -type NoteProperty -name $tmpName -value $tmpVal
                    }
                }
                break
            }
            '^_db_backups$' {
                # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                foreach ($val in $hldOptions.$version.dbBackupOptions.keys) {
                    $tmpName = $obj + "_" + $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.$version.dbBackupOptions[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")) {
                        $thisServer | add-member -type NoteProperty -name $tmpName -value $tmpVal
                    }
                }
                break
            }
            '^_db\d+_info$' {

                foreach ($val in $hldOptions.$version.dbOptions.keys) {
                    $tmpName = $obj + "_" + $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.$version.dbOptions[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")) {
                        $thisServer | add-member -type NoteProperty -name $tmpName -value $tmpVal 
                    }
                }
                break
            }
            '^_tsna_installed_software\d+$' {
                foreach ($val in $hldOptions.$version.tsnasoftware.keys) {
                    $tmpName = $obj + "_" + $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.$version.tsnasoftware[$val]
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                    if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")) {
                        $newSoftware += ,$tmpVal 
                    }
                }
                break
            }
            default {
                $tmpVal = $ws.cells.item($hldObjects[$obj], $col).Value()
                # this is for all other object requests

                if ((-not [string]::IsNullOrEmpty($tmpVal)) -and ($tmpVal -ne "(Choose)")) {
                    $thisServer | add-member -type NoteProperty -name $obj -value $tmpVal
                }
            }
        }
    }
    $ws = $null

    $correctedStorage = @($serverStorage[0])
    # Convert all hashes into pscustomobjects
    for ($i = 0; $i -lt ($serverStorage.count - 1); $i++) {

        if (($i -gt 0) -and ($serverstorage[$i].count -gt 1)) {
            $correctedStorage += ,(new-object -type PSCustomObject -property $serverStorage[$i])
        }
    }

    # Add installed software list to server
    $thisServer | add-member -type NoteProperty -name Software -value $newSoftware
    $thisServer | add-member -type NoteProperty -name Storage -value $correctedStorage

    # return $thisServer for inclusion in the FMO servers array
    $thisServer
}

#import CMO data from server definitions sheet in HLD
function Import-CMO ($cCol, $scount) {
    # import CMO data to $hld.CMO()
    $thisServer = new-object PSObject
    $thisServer | add-member -type NoteProperty -name _column -value $col
    $thisServer | add-member -type NoteProperty -name _disposition -value "CMO" | add-member -type noteproperty -name _hld_name -value $thishld.name
    $ocount = 0
    $col = $cCol

    # Create array of empty hashes
    $serverstorage = 0..15
    for ($i = 0; $i -lt 15; $i++) {
        # Initialize empty array of hashes
        $serverstorage[$i] = @{}
    }

    foreach ($obj in $hldObjects.keys | sort) {
        $ocount++
        if (-not $debug) { Write-Progress -Activity "Importing CMO data" -status "Server $scount of $($cmoColumns.count) - $obj" -percentComplete ((($scount * $ocount) / ($fmoColumns.count * $hldObjects.count)) * 100) }
        $debuginfo = @{}
        
        switch -regex ($obj) {
            '^_storage\d+$' {
                $blah = $obj -match '\D+(\d+)\D*'
                $objNum = $matches[1]
                # This is a storage object .. storage1, storage2 etc. We add the values for each $storage type to the fmohost column # to get the value
                foreach ($val in $hldOptions.CMOStorage.keys) {
                # Assign each value to $newserver hash keyed on object name
                    $tmpName = $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.CMOStorage.$val
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                    if (-not [string]::IsNullOrEmpty($tmpVal)) {
                            $serverStorage[$objNum][$tmpName] = $tmpVal
                    }
                }

                break
            }

            '^_storage_totals$' {
                # Total storage and migrating data
                $goodKeys = @("used", "free", "size")
                foreach ($val in $hldOptions.CMOStorage.keys) {
                    if ($goodkeys.contains($val)) {
                        $tmpName = $obj + "_" + $val
                        $tmpRow = $hldObjects[$obj]
                        $tmpCol = $col + $hldOptions.CMOStorage.$val
                        $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
                        if (-not [string]::IsNullOrEmpty($tmpVal)) {
                            $thisServer | add-member -type NoteProperty -name $tmpName -value $tmpVal 
                        }
                    }
                }
            }
            '^_db\d+_info$' {

                foreach ($val in $hldOptions.CMODbOptions.keys) {
                    $tmpName = $obj + "_" + $val
                    $tmpRow = $hldObjects[$obj]
                    $tmpCol = $col + $hldOptions.CMODbOptions.$val
                    $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()
    
                    if (-not [string]::IsNullOrEmpty($tmpVal)) {
                        $thisServer | add-member -type NoteProperty -name $tmpName -value $tmpVal 
                    }                      
                }
            }
            '^_monitoring_requirements' {
                # This is the start of the software section in CMO
                $cmoSoftware = @()
                # Just grab 15 rows of values
                for ($i = 0; $i -lt 15; $i++) {
                    $tmpRow = $hldObjects[$obj] + $i
                    $val = $ws.cells.item($tmpRow,$col).value()
                    if (-not [string]::IsNullOrEmpty($val)) {
                        $cmoSoftware += ,$val
                    }
                }

                $thisServer | add-member -type NoteProperty -name Software -value $cmoSoftware
            }
            default {
                $tmpVal = $ws.cells.item($hldObjects[$obj], $col).Value()
                if (-not [string]::IsNullOrEmpty($tmpVal)) {
                    $thisServer | add-member -type NoteProperty -name $obj -value $tmpVal
                }              
            }
        }
    }
    $ws = $null

    # Convert all hashes into pscustomobjects
    $correctedStorage = @($serverStorage[0])
    # Convert all hashes into pscustomobjects
    for ($i = 0; $i -lt ($serverStorage.count - 1); $i++) {
        if (($i -gt 0) -and ($serverstorage[$i].count -gt 1)) {
            $correctedStorage += ,(new-object -type PSCustomObject -property $serverStorage[$i])
        }
    }
    
    $thisServer | add-member -type NoteProperty -name Storage -value $correctedStorage
    # return $thisServer to be added to CMO array
    $thisServer
}

#import application requirements sheet 
function Import-AppReqs {
    # Import App requirements to $hld.Requirements{}
    $sheetname = "Application Requirements"
    $ws = $wb.worksheets | where {$_.name -eq $sheetName}
    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count
    $appreqObjects = @()
    $newObj = @{}

    $newObj["_overview"] = $ws.cells.item(2,2).value()

    for ($row = 4; $row -lt $rowCount; $row++) {            # Start at row4 to skip app overview
        $cellname = $ws.cells.item($row,1).value()
        if (-not $debug) { Write-Progress -Activity "Importing Application Requirements" -status "Object $row of $rowCount - $cellname" -percentComplete (($row/$rowcount) * 100) }
        if ($cellname -match '^_') {
            $cellval = $ws.cells.item($row,3).value()
            $cellnotes = $ws.cells.item($row,4).value()
            $newObj[$cellname] = @{"Value" = $cellval; "Notes" = $cellnotes }

        }
    }
    $appreqObjects += ,$newObj

    $ws = $null
    # Add $thisserver object to $thisHLD
    $thisHLD | add-member -type NoteProperty -name Requirements -value $appreqObjects -force
    $thisHLD
}

#import Overview sheet from HLD
function Import-Overview {
    $overviewObjs = @()
    $newobj = @{}
    $sheetname = "Overview"
    $ws = $wb.worksheets | where {$_.name -eq $sheetName}
    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count

    # Import overview sheet to $hld.Overview{}
    for ($row = 1; $row -lt $rowCount; $row++) { 
        $cellname = $ws.cells.item($row,1).value()
        if (-not $debug) { Write-Progress -Activity "Importing Overview" -status "Object $row of $($rowCount) - $cellname" -percentComplete (($row/$rowcount) * 100) }
        if ($cellname -match '^_') {
            $cellval = $ws.cells.item($row,6).value()
            $newobj[$cellname] = $cellval
        }
    }
    $overviewObjs += ,$newobj
    $ws = $null
    # Add $thisserver object to $thisHLD
    $thisHLD | add-member -type NoteProperty -name Overview -value $overviewObjs -force
    $thisHLD
}

<#
.SYNOPSIS
   Export HLD object to a new HLD template version v5.2+
.DESCRIPTION
   Exports HLD PSCustomObject containing all data to the specified HLD file

.EXAMPLE
    $hld = import-hld c:\path\to\hld\template
    $hld = import-hld c:\path\to\hld\template -force
    $hld = import-hld c:\path\to\hld\template -force -cmoOnly
    ls *.xlsm | import-hld -force -cmoOnly | export-hld -cmoOnly
      
#>
function global:Export-HLD  {
    [CmdletBinding()]
    param(
        [parameter (Mandatory=$true, ValueFromPipeline=$true, Position=0, ValueFromPipelineByPropertyName=$true)][Alias('MD5')][PSObject[]]$hlds,
        [parameter (Mandatory=$false, ValueFromPipeline=$false, Position=1, HelpMessage='Destination HLD file')][string]$outfile,
        [parameter (Mandatory=$false)][switch]$cmoOnly    
    )

    BEGIN {
    # Define global vars and load options definition file
    $templatePath = "$PSScriptRoot\CurrentTemplate.xlsm"
    $xl = new-object -comobject excel.application
    $xl.visible = $false
    $xl.displayalerts = $false
    $HLDOptions = import-clixml "$PSScriptRoot\options.hld"
    }
    PROCESS {
        $templatePath = read-host "Enter path to current HLD template"
        foreach ($hldObj in $hlds) {
            if (-not $outfile) {
                $outfile = "$($hldObj.name)_name"
            }

            # Copy template file to output file
            try {
                copy $templatePath $outfile
            }
            catch {
                write-error "Error copying template to destination path"
                break
            }
            
            # Open specified file
            try {
                $wb = $xl.Workbooks.Open($outfile)
            }
            catch {
                write-error "Error opening $outfile"
                break
            }

            # Write new overview tab to new template
            Export-OverviewTab
            # Write migration tab to new template
            Export-MigrationTab
            # Write server definitions tab to new template
            Export-ServerDefinitionsTab
            # Write application requirements tab to new template
            Export-ApplicationRequirementsTab

            # Save workbook with Overview tab active
            $sheetname = "Overview"
            $ws = $wb.worksheets | where {$_.name -eq $sheetName}
            $ws.Activate()

            # shut er down
            try {
                $wb.save()
                $wb.close()
            }
            catch {
                Write-error "Error saving file"
            }
        }
    }
    END {$xl.quit() }
}

#Export overview data to new template
function Export-OverviewTab {
    $sheetname = "Overview"
    $overviewObjs = @{}
    $ws = $wb.worksheets | where {$_.name -eq $sheetName}
    $ws.Activate()

    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count

    # index Overview sheet
    for ($row = 1; $row -lt $rowCount; $row++) { 
        $cellname = $ws.cells.item($row,1).value()

        # Validate format of object column and check ignore array before writing value
        if (($cellname -match '^_') -and (-not $hldOptions.Ignore.Contains($cellname))) {
            $ws.cells.item($row, 6).value() = $hldObj.Overview.$cellname
        }
    }
    $ws = $null
}

# populate migration tab with server info
function Export-MigrationTab {
    # Go to migration tab, and set cell 1,5 to num migration servers
    $sheetname = "Migration"
    $ws = $wb.worksheets | where {$_.name -eq $sheetname}
    $ws.Activate()
    $ws.cells.item(1,5).value() = $hldObj.CMO.count
    $startRow = 6

    # Populate migration tab
    for ($num = 0; $num -le $hldObj.CMO.count; $num++) {
        $ws.cells.item($startRow, 2).value() = $hldObj.CMO[$num]._fmo_hostname 
        $ws.cells.item($startRow, 3).value() = $hldObj.CMO[$num]._ip_cust 
        $ws.cells.item($startRow, 4).value() = $hldObj.CMO[$num]._description 
        $ws.cells.item($startRow, 5).value() = $hldObj.CMO[$num]._landscape
        $startRow++
    }

    $ws = $null
    sleep 10
}

# Populate app requirements in new HLD template
function Export-ApplicationRequirementsTab {
    # index Requirements sheet and insert data
    $sheetname = "Application Requirements"
    $ws = $wb.worksheets | where {$_.name -eq $sheetname}
    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count
    $ws.Activate()

    for ($row = 1; $row -lt $rowCount; $row++) { 
        $cellname = $ws.cells.item($row,1).value()
        # Insert data into app requirements tab
        if (($cellname -match '^_') -and (-not $hldOptions.ignore.contains($cellname))) {
            $ws.cells.item($row, 3).value() = $hldObj.Requirements.$cellname.Value
            $ws.cells.item($row, 4).value() = $hldObj.Requirements.$cellname.Notes   
        }
    }
    $ws = $null
}

# Populate server definitions tab
function Export-ServerDefinitionsTab {
    # Set view combobox to show all servers
    $sheetname = "Server Definitions"
    $ws = $wb.worksheets | where {$_.name -eq $sheetname}
    $ws.cells.item(300,1).Value() = 4

    #index server definitions tab
    $sheetname = "Server Definitions"
    $ws = $wb.worksheets | where {$_.name -eq $sheetname}
    $ws.Activate()

    # index server definitions tab    
    $usedRng = $ws.UsedRange.Cells
    $colCount = $usedRng.Columns.Count
    $rowCount = $usedRng.Rows.count

    # Index of server objects
    $serverObjs = @{}
    # Index of CMO column #s
    $cmoColumns = @()
    # Index of FMO Column #s
    $fmoColumns = @()

    # Index rows
    for ($row = 1; $row -lt $rowCount; $row++) { 
            $cellval = $ws.cells.item($row,1).value()
            if ($cellval -match '^_') {
                $serverObjs[$cellval] = $row 
            }
        }

    # Index sheet columns, and populate cmo/fmoColumn arrays with col number
    for ($col = 1; $col -lt $colCount; $col++) { 
        $cellval = $ws.cells.item(1,$col).value()
        if ($cellval -eq "FMO" -or $cellval -eq "New Demand") {
            $fmoColumns += ,$col
        } 
        elseif ($cellval -eq "CMO") {
            $cmoColumns += ,$col
        }
    }

    # insert cmo data from $hldobj into server definitions tab
    for ($cmo = 0; $cmo -lt $cmoColumns.count; $cmo++) {
        foreach ($sobj in $serverObjs.keys) {
            if ($sobj -match '^_storage(\d+)') {
                $num = $matches[1]
                # Write CMO storage 
                foreach ($key in $hldOptions.CMOStorage.keys) {
                    # Check ignore list
                    if (-not $hldOptions.Ignore.contains($key)) {
                        # Escape / for excel
                        $tempval = $hldObj.CMO[$cmo].Storage[($num)].$key
                        if (-not ([string]::IsNullOrEmpty($tempval)) -and ($tempval.gettype() -eq [string]) -and ($tempval.Startswith('/'))) {
                            $tempval = "'$tempval"
                        }
                        $tempCol = $cmoColumns[$cmo] + $hldOptions.CMOStorage.$key
                        $ws.cells.item($serverobjs[$sobj], $tempCol).value() = $tempval
                    }
                }
            }
            # Write CMO Software
            if ($sobj -match '^_monitoring_requirements') {
                $software = $hldObj.CMO[$cmo].Software -join ','
                $ws.cells.item($serverObjs[$sobj], $cmoColumns[$cmo]).value() = $software
            }
            else {
                $ws.cells.item($serverObjs[$sobj],$cmoColumns[$cmo]).value() = $hldObj.CMO[($cmo)].$sobj
            }
        }
    }

    if (-not $cmoOnly) {
        # Write FMO data to new HLD
        for ($fmo = 0; $fmo -lt $fmoColumns.count; $fmo++) {
            foreach ($sobj in $serverObjs.keys) {
                #if (-not $debug) { Write-Progress -Activity "Exporting server definitions" -status "Object $fmo of $($fmoColumns) - $obj" -percentComplete (($fmo/$fmoColumns.count) * 100) }
                
                # Skip if object is in ignore list
                if ($hldOptions.ignore.contains($sobj)) {
                    continue
                }
                switch -regex ($obj) {
                    '^_storage\d+$' {
                        # Retrieve the number from the object 
                        $obj -match '\D+(\d+)\D*'
                        $objNum = $matches[1]

                        foreach ($val in $hldOptions.$version.storage.keys) {
                            $tmpCol = $fmoColumns[$fmo] + $hldOptions.$version.Storage.$val
                            $newval = $hldObj.FMO[$fmo].Storage[$objNum].$val
                            $ws.cells.item($serverObjs[$sobj],$tmpCol).value() = $newval
                        }
                    } 
                    '^_storage\d+_details$' {

                        $obj -match '\D+(\d+)\D*'
                        $objNum = $matches[1]
                        # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                        foreach ($val in $hldOptions.$version.storageOptions.keys) {
                            $tmpCol = $fmoColumns[$fmo] + $hldOptions.$version.StorageOptions.$val
                            $newval = $hldObj.FMO[$fmo].Storage[$objNum].$sobj
                            $ws.cells.item($serverObjs[$sobj],$tmpCol).value() = $newval          
                        }
                    }<#
                    '^_storage_totals$' {
                        # Total storage and migrating data
                        foreach ($val in $hldOptions.$version.storageTotals.keys) {
                            $tmpCol = $fmoColumns[$fmo] + $hldOptions.$version.StorageTotals.$val
                            $newval = $hldObj.FMO.Storage[0].$sobj
                            $ws.cells.item($serverObjs[$sobj],$tmpCol).value() = $newval            
                        }
                    }
                    '^_fs_backups$' {
                        # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                        foreach ($val in $hldOptions.$version.fsBackupOptions.keys) {
                            $tmpCol = $fmoColumns[$fmo] + $hldOptions.$version.FSBackupOptions.$val
                            $newval = $hldObj.FMO.Storage[$fmo + 1].$sobj
                            $ws.cells.item($serverObjs[$sobj],$tmpCol).value() = $newval          
                        }
                    }
                    '^_db_backups$' {
                        # Storage_details1, storage_details2 etc. Add values from storage options to fmo hostname column to get the correct cell.
                        foreach ($val in $hldOptions.$version.dbBackupOptions.keys) {
                            $tmpCol = $fmoColumns[$fmo] + $hldOptions.$version.DBBackupOptions.$val
                            $newval = $hldObj.$serveNum.$sobj
                            $ws.cells.item($serverObjs[$sobj],$tmpCol).value() = $newval          
                        }
                    }
                    '^_db\d+_info$' {

                        foreach ($val in $hldOptions.$version.dbOptions.keys) {
                            $tmpCol = $fmoColumns[$fmo] + $hldOptions.$version.DBOptions.$val
                            $newval = $hldObj.$serveNum.$sobj
                            $ws.cells.item($serverObjs[$sobj],$tmpCol).value() = $newval         
                        }

                    }
                    '^_tsna_installed_software\d+$' {
                        # foreach ($val in $hldOptions.$version.tsnasoftware.keys) {
                        #     $tmpName = $obj + "_" + $val
                        #     $tmpRow = $hldObjects[$obj]
                        #     $tmpCol = $col + $tsnasoftware[$val]
                        #     $tmpVal = $ws.cells.item($tmpRow, $tmpCol).Value()

                        #     if ($tmpVal -ne '') { $newSoftware += ,$tmpVal }
                        # }
                    }#>
                    default {
                        $tmpcol = $fmoColumns[$fmo]
                        $newval = $hldObj.FMO[$fmo].$sobj

                        $ws.cells.item($serverObjs[$sobj],$fmoColumns[$fmo]).value() = $newval
                    }
                }
            }
        }
        $ws = $null
    }
}


function global:Prepare-HLDTemplate ($hldFile) {
    $xlShiftToRight = -4161

    # Arrays to hold property names
    $overview = @()
    $requirements = @()
    $serverDefs = @()

    # Open workbook
    $xl = new-object -comobject excel.application
    $xl.visible = $false
    $xl.displayalerts = $false
    $wb = $xl.Workbooks.Open($hldFile)

    # select Overview worksheet
    $ws = $wb.worksheets | where {$_.name -eq "Overview"}

    # Insert first column
    $range = $ws.range("A1").EntireColumn
    $range.insert($xlShiftToRight) | out-null

    # Write values to column1

    # Switch to server defs sheet

    # Insert first column
    $range = $ws.range("A1").EntireColumn
    $range.insert($xlShiftToRight) | out-null

    # write values to column 1

    # Switch to app requirements sheet

    # Insert first column
    $range = $ws.range("A1").EntireColumn
    $range.insert($xlShiftToRight) | out-null

    # Write values
}

<#
.SYNOPSIS
    Perform basic validation of HLD data
.DESCRIPTION
    Validates various fields of HLD object to ensure consistency

.EXAMPLE
    import-hld "c:\path to hld\file.xlsm" -force | validate-hld | export-hld -cmoOnly
#>
function global:Validate-HLD ($hldObj) {
    # validate fmo
    $count = 0
    foreach ($srv in $hldObj.FMO) {
        $count++
        # Check slice count
        if (($srv._ram * 2) -ne $srv._slices) {
            write-warning "Server $count slice count should be 2x ram"
        }

        # Check for windows support disks
        if ($srv._os_type -match '^W2K') {
            $tools = $false
            $pagefile = $false

            foreach ($storage in $srv.Storage) {
                if ($storage._usage -match 'Tools$') {
                    $tools = $true
                }
                if ($storage._usage -match 'Pagefile$') {
                    $pagefile = $true
                }
            }
            if (-not $tools) {
                write-warning "Server $count TSNA Tools drive not found"
            }
            if (-not $pagefile) {
                write-warning "Server $count pagefile drive not found"
            }
        }

        # Check OS edition for ram
        if ($srv._os_type -match '^W2K\d+SE') {
            if ($srv._ram -gt 32) {
                write-warning "Server $count requires Enterprise OS for full RAM access"
            }
        }

        # Validate SQL is defined if in CMO
        if ($srv._db_vendor -eq 'No Change') {
            write-warning "Server $count - SQL version should be set to $($hldObj.$CMO[$count]._db_vendor)"
        }

        # Check for greenfield user access
        if ($srv._migration_type -eq 'Greenfield') {
            if ([string]::IsStringNullOrEmpty($srv._user_access)) {
                write-warning "Server $count is missing greenfield usernames"
            }
        }
    }
}

<#
.SYNOPSIS
    Basic reporting of HLD data
.DESCRIPTION
    format-hld HLD data into standard format for account review.

.EXAMPLE
    format-hld $hld c:\user\some\output.csv.
#>
function global:Format-HLD {
    # Requirements from Rhett
    # 1.  HLD name
    # 2.  CMO server name
    # 3.  FMO server name
    # 4.  OS version
    # 5.  Database version, if exists
    # 6.  Cores
    # 7.  Memory
    # 8.  Storage tier
    # 9.  Storage config (tricky, because different servers have different layouts of course): drive/mount point and size in GB
   #$hldObj.fmo | select _hld_name, _cmo_hostname,_fmo_hostname,_migration_type,_platform,_slices,_cpu_cores,_ram,_db_vendor -expandproperty Storage
    #$hld.fmo | select -property _hld_name, _fmo_hostname, _cmo_hostname, _migration_type, _platform, _slices, _cpu_cores, _ram, _db_vendor, @{Name="Drive"; Expression={for ($i = 1; $i -lt $_.Storage.count; $i++) {$_.Storage[$i].ID} }}, @{Name = "Tier"; Expression={for ($i = 1; $i -lt $_.Storage.Count; $i++) { $_.Storage[$i].tier}}}, @{Name="Size"; Expression={for ($i = 1; $i -lt $_.Storage.count; $i++) {$_.Storage[$i].size}}}, @{Name="Total Storage"; Expression={$_.Storage[0].total}} | Export-Csv 'C:\users\nlucent\documents\spideroak hive\work stuff\test.csv'
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$false)][string]$hldpath,
        [parameter(Mandatory=$true, ValueFromPipeline=$false)][string]$outfile
    )

    $ci = Get-ChildItem $hldpath

    $hld = import-hld $ci.FullName 


    foreach ($s in $hld.FMO) {
        $srv = $s | select-object _hld_name,_cpu_cores,_criticality,_db_vendor,_fmo_dc,_fmo_hostname,_landscape,_migration_type,_os_type,_platform,_priority,_ram, _server_type,_slices
        for ($i = 1; $i -lt $s.Storage.count; $i++) {
            $sname = "Storage" + $i
            $srv | add-member -type NoteProperty -name $($sname + "_Usage") -value $s.Storage[$i].Usage
            $srv | add-member -type NoteProperty -name $($sname + "_Options") -value $s.Storage[$i].Options
            $srv | add-member -type NoteProperty -name $($sname + "_ID") -value $s.Storage[$i].ID
            $srv | add-member -type NoteProperty -name $($sname + "_Type") -value $s.Storage[$i].Type
            $srv | add-member -type NoteProperty -name $($sname + "_Size") -value $s.Storage[$i].Size
            $srv | add-member -type NoteProperty -name $($sname + "_Format") -value $s.Storage[$i].Format
            $srv | add-member -type NoteProperty -name $($sname + "_BackupType") -value $s.Storage[$i].BackupType
            $srv | add-member -type NoteProperty -name $($sname + "_Multiplier") -value $s.Storage[$i].Multiplier
            $srv | add-member -type NoteProperty -name $($sname + "_Tier") -value $s.Storage[$i].Tier
            $srv | add-member -type NoteProperty -name $($sname + "_DataToMove") -value $s.Storage[$i].DataToMove

        }

        if (test-path $outfile) {
            $srv | export-csv -path $outfile -append -force -noTypeInformation
        }
        else {
            $srv | export-csv -path $outfile -force -noTypeInformation
        }
    }
}

# From the excel cookbook, may not be necessary.
function Remove-ComObject {
    # Requires -Version 2.0
    [CmdletBinding()]
    param()
    end {
        Start-Sleep -Milliseconds 500
        [Management.Automation.ScopedItemOptions]$scopedOpt = 'ReadOnly, Constant'
        Get-Variable -Scope 1 | Where-Object {
            $_.Value.pstypenames -contains 'System.__ComObject' -and -not ($scopedOpt -band $_.Options)
            } | Remove-Variable -Scope 1 -Verbose:([Bool]$PSBoundParameters['Verbose'].IsPresent) | out-null
            [gc]::Collect() | out-null
    }
}

# Export functions
export-modulemember -function Import-HLD, Export-HLD, Format-HLD
export-modulemember -function Prepare-HLDTemplate, Validate-HLD
