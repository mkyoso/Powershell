#########################################################
### ファイル比較処理
<#
    本スクリプトはメインファイルから呼び出され実行されます。
    本処理では"Compare-Object"を使用しテキストファイルの比較差分の出力処理を実施します。
    "Compare-Object"では、差分表現が以下となっており分かりにくいため表現を変更する処理となっています。

　　 ●Compare-Object 比較表現
    ・同一 … "=="
    ・追加/変更 … "=>"
    ・削除 … "<="

#>

function Global:act_diffs{
    param(
        [String]$FromFile,
        [String]$ToFile
    )
    $result = @()
    Compare-Object (Get-Content $FromFile) (Get-Content $ToFile) | Select-Object -Property @{Name = "ReadCount"; Expression = {$_.InputObject.ReadCount}}, * | Sort-Object -Property ReadCount |
        ForEach-Object {
            [string]$line = ""
            [string]$foreColor = ""
            if ($_.SideIndicator -eq "=>")
            {
                # 修正後に存在する行（追加または変更された行）
                $line = "[+] " + $_.InputObject + "`n"
                $foreColor = "Green"
            }
            elseif ($_.SideIndicator -eq "<=")
            {
                # 修正後に存在しない行（削除または変更された行）
                $line = "[-] " + $_.InputObject + "`n"
                $foreColor = "Magenta"
            }
            elseif ($Full)
            {
                # 変更がない行
                $line = "[=] " + $_.InputObject + "`n"
                $foreColor = "Gray"
            }
            $result += $line
        }
        $line | Out-File $out_fs
        return $result
}
