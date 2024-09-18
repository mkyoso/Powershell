#########################################################
### ファイル(CSV/Excel)整形用スクリプト
#########################################################
<#
　本処理はファイル(CSV/Excel)整形処理ツールとなります。

　■ 更新履歴
　202403dd ... 初版(テスト中)


  ◆実行時セキュリティエラーにて実行できない場合
  　本スクリプト実行時、以下のセキュリティエラーが発生する場合は以下の処理を実行前に実施してください。

  　1.エラー内容
  　　このシステムではスクリプトの実行が無効になっているため、ファイル ～domo_uploder_v2.ps1 を読み込むことができません。
  　　詳細については、「about_Execution_Policies」(https://go.microsoft.com/fwlink/?LinkID=135170) を参照してください。
    　　　+ CategoryInfo          : セキュリティ エラー: (: ) []、ParentContainsErrorRecordException
    　　　+ FullyQualifiedErrorId : UnauthorizedAccess
  　・対処方法
  　　本エラーはPowerShell実行権限がないため出力されます。
  　　一時解決として、以下のコマンドレットを実行することで実行権限を付与することが可能です。
  　　※ Windowsの既定値ではPowerShell実行権限は許可されていません。
  　　
  　　●確認/対処方法
  　　　# 現状のセキュリティポリシー確認
  　　　> Get-ExecutionPolicy
  　　　Restricted　#既定値
  　　　# Powershell実行権限の付与
        > Set-ExecutionPolicy Remotesigned -Scope Process
   　　 > Get-ExecutionPolicy
 　　　 Remotesigned #実行権限が付与されたことを確認

  　2.エラー内容
  　　ネットワークドライブ内ファイルへの参照が正常に行えない場合

  　・対処方法
  　　本エラーはSharePointのWEB共有マッピングの認証資格情報が切れている、認証セッションの有効期限切れで発生します。
  　　※ Windows HTTPサービス(WinHTTP)を使用していることで発生する事象となります。
         (Microsoft 365アカウントを切り替える等によってもセッションが切れる可能性があります)
  　　一時解決として、以下の操作を行うことによりドライブマッピングの復旧が可能です。
  　　
  　　●対処方法
  　　　1) Webブラウザより共有ネットワークドライブ対象のSharePointOnlineドキュメントライブラリを開く
  　　　2) SharePointOnlineドキュメントライブラリ内、上部メニュー[すべてのドキュメント]->[エクスプローラーで表示]を選択
    　　3) 実行端末のエクスプローラーにて該当ドキュメントライブラリが表示されたことを確認し、エクスプローラー内[PC]へ移動
    　　4) 上部メニュー[ネットワークドライブの割り当て]より共有ネットワークドライブ対象のSharePointOnlineドキュメントライブラリを設定

    #.補足

#>
#########################################################
# パラメーター定義
<#

#>
Param()

#########################################################
# 変数定義
<#
　
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"

#########################################################
# 変数定義

## 基本設定
$yyyymmdd = Get-Date -Format "yyyyMMdd"
$root_path = $PSScriptRoot
$Log_base = $root_path + "\Logs\"
$Log_path = $Log_base + "log_" + $yyyymmdd + ".log"
#$conf_base = $root_path + "\configs\"
#$conf_path = $conf_base + "config.ini"

## 実行時の年度取得
$FY = (Get-Date).AddMonths(-3).ToString("yyyy")

## 実行時の年月取得
# スクリプト実行時の年を取得
$YEAR = Get-Date -Format yyyy
# スクリプト実行時の月を取得
$org_mon = Get-Date -Format MM
# スクリプト実行時の月(短縮形：M)を取得
$SMONTH = (Get-Date).Month
# ファイルへ追記する利用月(先月)を設定
$MONTH = ($org_mon -2).ToString("00") # 検証用：$org_mon -2、本番用：$org_mon -1

## 実行時の先月取得
$PREV_MONTH = (Get-Date).AddMonths(-1).ToString("MM")

#########################################################
# 実行処理

## 事前処理
# Log取得開始
Start-Transcript $Log_path -Append
Write-Verbose "スクリプトツール ... Start"

# Functions配下関数ファイル読み込み
$read_func = Get-ChildItem "${root_path}\functions" -Recurse -Filter *.ps1
$read_md = "${root_path}\module"
Write-Host "外部関数呼び出し処理...Start"
foreach($func in $read_func){
    Write-Host "${func}ファイルを読み込みます"
    . $func.FullName
}
Import-Module "${root_path}\module\ImportExcel.psm1"
cm_msg "外部関数呼び出し処理...Success"
# configs\config.ini読み込み
cm_header "Config呼び出し処理"
#$ini =@{}
#Get-Content $conf_path | Where-Object {$_ -notmatch '^\s*$'} | Where-Object {!($_.TrimStart().StartsWith(";"))}

## 変数設定
<#
  変数を指定して下さい
  以下はサンプル変数となります
#>
# ここから
# 入力ファイル用変数
$SRC_FILE1 = "_" + $YEAR + $SMONTH + ".csv"
$SRC_FILE2 = "_" + $YEAR + $SMONTH + ".csv"
$SRC_XLSX_FILE1 = "test.xlsx"
$SRC_XLSX_FILE1_PATH = $root_path + "\src\"

# 出力ファイル用変数
$DEST = $root_path + "\out\"
$ADD_FS1 = "test_csv_1.csv"
$ADD_FS2 = "test_csv_1.csv"
$ADD_FS3 = "test_xlsx_1.xlsx"
# ここまで
cm_footer "Config呼び出し処理"

<# 主処理
実行処理となります。
実行される処理は以下の通り。

    1. コピー先フォルダ内削除処理
        コピー元ファイル送り先のフォルダ内に格納されている全ファイルの削除処理
    2.ファイル整形処理(判断分岐処理)
    　ファイルパスおよびファイル名を元に以下の処理を実施
        2.1. CSVファイル カラム追加処理(例：年月)
            本処理は外部関数にて処理を実施
            カラム追加後のファイルはコピー先に生成
            整形処理後、Diff比較を実施

        2.2. Excelファイル カラム追加処理(例：年月)
            本処理は変数定義にて定義しているExcelオブジェクトを活用し、対象ワークシート内の最終行・列を取得して最終列+1に年月の追加を実施
            カラム追加後のファイルはコピー先に生成

#>

# コピー先フォルダ内削除処理
cm_header "コピー先フォルダ内削除処理"
cm_msg "コピー先フォルダ一覧確認処理...Start"
if(Test-Path $DEST){
    $ck_dst_fs = Get-ChildItem $DEST -File
    if($ck_dst_fs){
        $ck_df = $ck_dst_fs.Name
        $msg = $ck_df + " ... 削除対象ファイル存在"
        cm_msg $msg
        # フォルダ内ファイル削除処理
        cm_msg "コピー先フォルダ一内削除処理...Start"
        foreach( $del_fs in $ck_dst_fs ){
            cm_msg $del_fs + "...削除開始"
            $del_fs_name = $del_fs.FullName
            Remove-Item $del_fs_name -Force
            cm_msg $del_fs + "...削除"
        }
    }
    else{
        $msg = "削除対象ファイル:なし"
        cm_msg $msg
    }
    $ck_dst_del = Get-ChildItem $DEST -File
    cm_msg $ck_dst_del
    cm_footer "コピー先フォルダ削除処理"
}
else{
    cm_msg "コピー先フォルダ一覧確認処理...Error"
}
cm_msg "コピー先フォルダ一覧確認処理...Success"

# ファイル整形処理
#$params = @( 変数で設定したパスを入力 )
$params = @(  $SRC_XLSX_FILE1_PATH )
cm_header "ファイル整形処理"
foreach($p in $params){
    $msg = "フォルダ確認処理:" + $p
    cm_msg $msg
    # ※接続ドライブへ移動(ネットワークドライブ(Sharepointドキュメント)を参照する場合の対策処理)
    # 移動せずに実行した場合、ドライブレターを正常に判別できません。
    #   (実態は「\\～.sharepoint.com～」というUNCパスとなりOSドライブレターはエイリアスとなります。)
    # ■実行時コメントアウト
    #cd X:
    if(Test-Path $p){
        cm_msg "フォルダ確認処理...Success"
        $ck_f = Get-ChildItem $p -File
        foreach($f in $ck_f){
            $ck_f_path =  $f.FullName
            $ck_fs = $f.Name
            <# CSVファイル処理
                CSVファイルの整形を実施します。
                CSV整形処理は外部関数(act_csv.ps1)にて実施されます。
                外部関数からの戻り値はなく、外部関数にて出力処理が行われます。

                # 処理工程
                    1.入力ファイル名検索
                    ①入力ファイル名が変数指定ファイル名と同一の場合
                        ・整形処理開始(cm_msg呼び出し)
                        ・出力ファイル存在確認
                        1)出力ファイルが存在する場合
                            処理スキップ
                        2)出力ファイルが存在しない場合
                            CSV整形関数呼び出し(act_csv.ps1)
                            # 指定変数
                            ・$src_fs ... 入力ファイルのフルパス
                            ・$after_fs ... 出力ファイルのフルパス
                            ・$act ... 整形内容
                            ②入力ファイル名が変数指定ファイル名と同一でない場合
                            スキップ
            #>
            if($ck_fs -eq  $SRC_FILE1){
                $s_msg = $ck_fs + " ... 整形処理"
                cm_msg $s_msg
                $dt = $YEAR + $MONTH
                $after_fs = $DEST + $ADD_FS1
                $src_fs = $ck_f_path
                if(Test-Path $after_fs){
                    $c_msg = $ck_f.Name + " ... 整形処理済み"
                    cm_msg $c_msg
                }
                else{
                    act_csv $src_fs $after_fs $dt
                    $e_msg = $ck_f.Name + " ... 整形処理完了"
                    cm_msg $e_msg
                }
            }
            elseif($ck_fs -eq  $SRC_FILE2){
                $s_msg = $ck_fs + " ... 整形処理"
                cm_msg $s_msg
                $dt = $YEAR + $MONTH
                $after_fs = $DEST + $ADD_FS2
                $src_fs = $ck_f_path
                if(Test-Path $after_fs){
                    $c_msg = $ck_f.Name + " ... 整形処理済み"
                    cm_msg $c_msg
                }
                else{
                    act_csv $src_fs $after_fs $dt
                    $e_msg = $ck_f.Name + " ... 整形処理完了"
                    cm_msg $e_msg
                }
            }
            <#
                対象CSVファイルが複数ある場合は上述の"elseif"をコピーペーストしてください。
                対象CSVが一つの場合は上述の"elseif"を削除してください。
            #>
            elseif($ck_fs -eq $SRC_XLSX_FILE1){
            <# Excelファイル整形処理
                Excelファイルの整形を実施します。
                Excelファイル整形処理はCOMオブジェクトを利用して実施されます。
                外部関数からの戻り値はなく、外部関数にて出力処理が行われます。

                # 処理工程
                    1.入力ファイル名検索
                        ①入力ファイル名が変数指定ファイル名と同一の場合
                        ・出力ファイル存在確認
                            1)出力ファイルが存在する場合
                                処理スキップ
                            2)出力ファイルが存在しない場合
                                1.対象ファイルのコピー処理
                                1)出力ファイルが存在する場合
                                    処理スキップ
                                2)出力ファイルが存在しない場合
                                    1.Excelアプリケーションオブジェクトの読み込み
                                    2.Excelワークシート最終列に追記処理
                                    3.Excelファイルの保存
                                    4.ExcelファイルDiff
                                    ②入力ファイル名が変数指定ファイル名と同一でない場合
                                    スキップ
            #>
                # 処理内変数
                $xlsx_file_path = $SRC_XLSX_FILE1_PATH + $SRC_XLSX_FILE1
                $dst_fs = $DEST + $ADD_FS3
                $YYMM = $YEAR + $MONTH
                if(Test-Path $dst_fs){
                    $c_msg = $dst_fs + " ... 整形処理済み"
                    cm_msg $c_msg
                }
                else{
                    $c_msg = $xlsx_file_path + " ...Excelファイルオープン処理"
                    cm_msg $c_msg
                    ## Excelファイル整形：事前処理
                    Add-Type -Path "${root_path}\module\EPPlus.dll"
                    # ファイルコピー処理
                    $src_xlsx = Import-Excel $xlsx_file_path
                    # Excelファイル整形処理
                    if($src_xlsx){
                        ## Excelオブジェクト
                        # Excelファイルオープン処理
                        $excel = Import-Excel $xlsx_file_path
                        # Excel Key/Valueオブジェクト追加
                        $excel | Add-Member NoteProperty -Name "日付" -Value $YYMM

                        # Excelファイル保存処理
                        $excel | Export-Excel $dst_fs

                        # ファイルチェック
                        $ck_fs=  Compare-Object $xlsx_file_path $dst_fs -IncludeEqual
                        cm_msg $ck_fs
                    }
                    else{
                        $e_msg = "Excelファイル操作エラー"
                        cm_msg $e_msg
                    }
                }
            }
        }
    }
    else{
        $p
        cm_msg "ファイル整形処理:フォルダ確認...Error"
    }
}
cm_footer "ファイル整形処理"

## 事後処理
# 生成ファイル確認処理
<# 事後処理
    本処理ではファイル整形後の事後処理として出力ディレクトリ内に生成ファイル(出力ファイル)が存在するかを確認します。
    elseif構文は確認対象ファイル数毎にコピー追加して。
#>
cm_header "生成ファイル確認処理"
cm_msg "生成ファイル確認処理...Start"
if(Test-Path $DEST){
    $ck_dst_dir = Get-ChildItem $DEST -File
    # フォルダ内ファイル確認処理
    foreach( $create_fs in $ck_dst_dir ){
        $msg = $create_fs.Name + " ... 生成ファイル確認"
        if($create_fs.Name -eq $ADD_FS1){
            cm_msg $msg
        }
        elseif($create_fs.Name -eq $ADD_FS3){
            cm_msg $msg
        }
    }
}
cm_footer "生成ファイル確認処理"

Stop-Transcript
