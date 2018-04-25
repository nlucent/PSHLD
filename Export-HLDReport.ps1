param(
    [parameter(Mandatory=$true, ValueFromPipeline=$false)][string]$hldpath,
    [parameter(Mandatory=$true, ValueFromPipeline=$false)][string]$outfile
    )

import-module .\pshld

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
