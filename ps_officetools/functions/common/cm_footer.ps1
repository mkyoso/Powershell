function Global:cm_footer{
    $msg_ary = @("�I������:$(Get-Date)","===================================================")
    foreach($msg in $msg_ary){
        Write-Host $msg
        $logger.info($msg)
    }
}