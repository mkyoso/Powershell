#########################################################
### CSVファイル整形/出力処理
<#
　■ 本スクリプトはメインファイルから実行されます。
　　本処理では"Import-Csv"にてCSVファイルを読み込み「Out-GridView」にて出力する処理を実施します。
#>

function Global:act_csv{
    param(
        [String]$fs_name
    )
    # 処理変数
    try{
        # CSVファイル読み込み
        $csv = Import-CSV $fs_name
        # GridView(GUI) 表示
        $csv  | Out-GridView -OutputMode Multiple
    }
    catch{
        $err = $error[0] | Format-List -force
        return $err
    }
}