#########################################################
### ログメッセージ出力処理
<#
　
#>

function Global:cm_msg{
    param(
        $message
    )
    Write-Host $message
}