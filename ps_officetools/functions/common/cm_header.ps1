function Global:cm_header{
    Param(
        [String]
        $Action
    )
    $msg_ary = @("===================================================","$($Action)�J�n","�J�n����:$(Get-Date)","===================================================")
    foreach($msg in $msg_ary){
        Write-Host $msg
        $logger.info($msg)
    }
}