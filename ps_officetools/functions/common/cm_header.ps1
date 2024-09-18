function Global:cm_header{
    Param(
        [String]
        $Action
    )
    $msg_ary = @("===================================================","$($Action)開始","開始時刻:$(Get-Date)","===================================================")
    foreach($msg in $msg_ary){
        Write-Host $msg
        $logger.info($msg)
    }
}