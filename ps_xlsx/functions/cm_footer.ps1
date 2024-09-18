#########################################################
### ログ処理ヘッダー出力処理
<#
　
#>

function Global:cm_footer{
    param(
        [String]
        $message
    )
    $msg_ary = @(
        "--------------------",
        "$($message):Stop",
        "StopTime:$(Get-Date)",
        "====================="
    )
    foreach($msg in $msg_ary){
        Write-Host $msg
    }
}