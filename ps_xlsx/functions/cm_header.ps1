#########################################################
### ログ処理ヘッダー出力処理
<#
　
#>

function Global:cm_header{
    param(
        [String]
        $message
    )
    $msg_ary = @(
        "======================",
        "$($message):Start",
        "StartTime:$(Get-Date)",
        "----------------------"
    )
    foreach($msg in $msg_ary){
        Write-Host $msg
    }
}