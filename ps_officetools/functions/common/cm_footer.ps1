function Global:cm_footer{
    $msg_ary = @("èIóπéûçè:$(Get-Date)","===================================================")
    foreach($msg in $msg_ary){
        Write-Host $msg
        $logger.info($msg)
    }
}