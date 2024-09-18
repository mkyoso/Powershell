function Global:cm_checkprocess{
    param(
        [Parameter()]
        [Switch]
        $outlook,
        [Parameter()]
        [Switch]
        $excel
    )
    $title = "Office�I�u�W�F�N�g�v���Z�X�m�F����"
    cm_header $title
    if($outlook){
        $logger.info("Outlook�m�F����")
        $check =Get-Process|Where-Object {$_.Name -match "OUTLOOK"}
        if ($check -eq $null){
            $existsOutlook = $false
            return $existsOutlook
        }
        else{
            $existsOutlook = $true
            return $existsOutlook
        }
    }
    elseif($excel){
        $logger.info("Excel�m�F����")
        $check =Get-Process|Where-Object {$_.Name -match "EXCEL"}
        if ($check -eq $null){
            $existsOutlook = $false
            return $existsOutlook
        }
        else{
            $existsOutlook = $true
            return $existsOutlook
        }
    }
}