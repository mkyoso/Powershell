function Global:cm_checkprocess{
    param(
        [Parameter()]
        [Switch]
        $outlook,
        [Parameter()]
        [Switch]
        $excel
    )
    $title = "Officeオブジェクトプロセス確認処理"
    cm_header $title
    if($outlook){
        $logger.info("Outlook確認処理")
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
        $logger.info("Excel確認処理")
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