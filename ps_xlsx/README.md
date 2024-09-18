# テキストファイル(CSV/xlsx[Excel])整形支援 PowerShell

## ファイル取得/CSV データ整形・出力処理スクリプトサンプル

---

本スクリプトは CSV・Excel(xlsx)データ整形出力スクリプトとなります。
実行時、Powershell セキュリティポリシー等にて実行できない場合は以下実行方式で使用が可能です。

> ```powershell
> powershell -ExecutionPolicy RemoteSigned ./{ファイル名}.ps1
> ```

本スクリプトの想定は前月分の CSV・Excel 情報を取得し、必要となるカラムを新規に追加。
追加カラムへ特定情報を追記したファイルを生成する処理となります。

---

### 利用方法

1)適当なディレクトリに本スクリプト一式を解凍します。
1-1)「{ファイル名：標準ファイル名‗ps_xlsx_formating}.ps1」を右クリックしメニューより"PowerShell で実行"。
2-1)権限により実行できない場合、以下コマンドプロンプト実行を実施。

> ```powershell
> powershell -ExecutionPolicy RemoteSigned ./ps_CSV_Shaping.ps1
> ```

3)本スクリプト内で利用する場合「src」フォルダ内にデータを格納。
　本スクリプト内以下ブロックを編集し、整形対象 CSV ファイルの変数定義を実施

> ```powershell
> ## 変数設定
> <#
>  変数を指定して下さい
>  以下はサンプル変数となります
> #>
> # ここから
> # 入力ファイル用変数
> $SRC_FILE1 = "_" + $YEAR + $SMONTH + ".csv"
> $SRC_FILE2 = "_" + $YEAR + $SMONTH + ".csv"
> $SRC_XLSX_FILE1 = "test.xlsx"
> $SRC_XLSX_FILE1_PATH = $root_path + "\src\"
>
> # 出力ファイル用変数
> $DEST = $root_path + "\out\"
> $ADD_FS1 = "test_csv_1.csv"
> $ADD_FS2 = "test_csv_1.csv"
> $ADD_FS3 = "test_xlsx_1.xlsx"
> # ここまで
> ```

4)本スクリプトを実行し、「out」フォルダ内に整形済みファイル一式が出力されていることを確認

###　スクリプト構成
本スクリプトのフォルダ・ファイル構成は以下となります。

📦ps_xlsx  
┣ 📂configs  
┣ 📂functions  
┣ 📂Logs  
┣ 📂module  
┃ ┣ 📂Charting  
┃ ┣ 📂en  
┃ ┣ 📂Examples  
┃ ┃ ┣ 📂AddImage  
┃ ┃ ┣ 📂AddWorkSheet  
┃ ┃ ┣ 📂Charts  
┃ ┃ ┣ 📂CommunityContributions  
┃ ┃ ┣ 📂ConditionalFormatting  
┃ ┃ ┣ 📂ConvertFrom  
┃ ┃ ┣ 📂CustomizeExportExcel  
┃ ┃ ┣ 📂CustomNumbers  
┃ ┃ ┣ 📂CustomReporting  
┃ ┃ ┣ 📂ExcelBuiltIns  
┃ ┃ ┣ 📂ExcelDataValidation  
┃ ┃ ┣ 📂ExcelToSQLInsert  
┃ ┃ ┣ 📂Experimental  
┃ ┃ ┣ 📂Extra  
┃ ┃ ┣ 📂Fibonacci  
┃ ┃ ┣ 📂FormatCellStyles  
┃ ┃ ┣ 📂FormatResults  
┃ ┃ ┣ 📂Freeze  
┃ ┃ ┣ 📂GenerateData  
┃ ┃ ┣ 📂Grouping  
┃ ┃ ┣ 📂HeaderName  
┃ ┃ ┣ 📂HyperLinks  
┃ ┃ ┣ 📂Import-Excel  
┃ ┃ ┣ 📂ImportByColumns  
┃ ┃ ┣ 📂ImportColumns  
┃ ┃ ┣ 📂ImportHtml  
┃ ┃ ┣ 📂InteractWithOtherModules  
┃ ┃ ┃ ┣ 📂Pester  
┃ ┃ ┃ ┗ 📂ScriptAnalyzer  
┃ ┃ ┣ 📂InvokeExcelQuery  
┃ ┃ ┣ 📂JoinWorksheet  
┃ ┃ ┣ 📂JustCharts  
┃ ┃ ┣ 📂MergeWorkSheet  
┃ ┃ ┣ 📂MortgageCalculator  
┃ ┃ ┣ 📂MoveSheets  
┃ ┃ ┣ 📂MultiplierTable  
┃ ┃ ┣ 📂Nasa  
┃ ┃ ┣ 📂NumberFormat  
┃ ┃ ┣ 📂OpenExcelPackage  
┃ ┃ ┣ 📂OutTabulator  
┃ ┃ ┣ 📂PassThru  
┃ ┃ ┣ 📂PesterTestReport  
┃ ┃ ┣ 📂PivotTable  
┃ ┃ ┣ 📂PivotTableFilters  
┃ ┃ ┣ 📂Plot  
┃ ┃ ┣ 📂ReadAllSheets  
┃ ┃ ┣ 📂SetColumnBackgroundColor  
┃ ┃ ┣ 📂Sparklines  
┃ ┃ ┣ 📂SpreadsheetCells  
┃ ┃ ┣ 📂SQL+FillColumns+Pivot  
┃ ┃ ┣ 📂Stocks  
┃ ┃ ┣ 📂Styles  
┃ ┃ ┣ 📂Tables  
┃ ┃ ┣ 📂TestRestAPI  
┃ ┃ ┣ 📂VBA  
┃ ┃ ┣ 📂XlRangeToImage  
┃ ┣ 📂FAQ  
┃ ┣ 📂InferData  
┃ ┣ 📂package  
┃ ┃ ┗ 📂services  
┃ ┃ ┗ 📂metadata  
┃ ┃ ┗ 📂core-properties  
┃ ┃ ┗ 📜b71ffc68c5d34ffd887454f84d7b9b36.psmdcp  
┃ ┣ 📂Pivot  
┃ ┣ 📂Private  
┃ ┣ 📂Public  
┃ ┣ 📂Testimonials  
┃ ┣ 📂_rels  
┃ ┣ 📜EPPlus.dll  
┣ 📂out  
┣ 📂src  
┣ 📜ps_xlsx_formating.ps1  
┗ 📜README.md  

「config」内、設定ファイルについては未実装となります。

#### 各ファンクション/処理ブロック説明

---

##### common ファンクション

1. cm_footer
   コンソールログ・ログファイル出力時に処理単位でのフッターを追記処理

> ```powershell
> function Global:cm_footer{
>     param(
>         [String]
>         $message
>     )
>     $msg_ary = @(
>         "--------------------",
>         "$($message):Stop",
>         "StopTime:$(Get-Date)",
>         "====================="
>     )
>     foreach($msg in $msg_ary){
>         Write-Host $msg
>     }
> }
> ```

2. cm_header
   コンソールログ・ログファイル出力時に処理単位でのヘッダーを追記処理します。

> ```powershell
> function Global:cm_header{
>    param(
>        [String]
>        $message
>    )
>    $msg_ary = @(
>        "======================",
>        "$($message):Start",
>        "StartTime:$(Get-Date)",
>        "----------------------"
>    )
>    foreach($msg in $msg_ary){
>        Write-Host $msg
>    }
> }
> ```

3. cm_msg
   コンソールログへのメッセージ追記処理

> ```powershell
>    Write-Host $message
> ```

共通処理は他スクリプト作成等にも流用が可能な様に簡素にしています。
logger を生成される場合については log4net を使用するなどして下さい。

---

##### action ファンクション

1. act_csv
   本処理では"Import-Csv"ではなく"Get-Content"によりテキストファイル整形処理を実施します。
   "Import-Csv"ではない理由として対象ファイルにヘッダーが無いためデータオブジェクトの格納が正常に行えないことからテキスト整形処理としています。
   また、整形後ファイルを事前に作成しているのは"Get-Content"系処理が Read Lock する関係上、複数処理中にプロセス重複が発生することを抑止するためとなります。

> ```powershell
> function Global:act_csv{
>     param(
>         [String]
>         $fs_name,
>         $after_fs,
>         $dt
>     )
>     # 処理変数
>     $dir = $after_fs
>     $result =@()
>     ## 実行部
>     # 作業ドライブへ移動(ネットワークドライブパス対策)
>     cd X:
>
>     # 整形ファイル生成
>     New-item $dir -ItemType File
>     # 整形元ファイル読み込み+整形ファイル出力
>     $data = Get-Content $fs_name -Encoding Default | % {$_ + ",${dt}"} | Out-File $dir -Append -Encoding default
>
>     ## ファイル比較処理
>     #  整形前ファイルと整形後ファイルのDiff出力処理
>     Compare-Object -ReferenceObject @(Get-Content $fs_name) -DifferenceObject @(Get-Content $dir) | Select-Object -Property @{Name = "ReadCount"; Expression = {$_.InputObject.ReadCount}}, * | Sort-Object -Property ReadCount
> }
> ```

2. Excel ファイル整形処理
   本処理はメインプロセスより実行されます。

> ```powershell
>                # 処理内変数
>                $xlsx_file_path = $SRC_XLSX_FILE1_PATH + $SRC_XLSX_FILE1
>                $dst_fs = $DEST + $ADD_FS3
>                $YYMM = $YEAR + $MONTH
>                if(Test-Path $dst_fs){
>                    $c_msg = $dst_fs + " ... 整形処理済み"
>                    cm_msg $c_msg
>                }
>                else{
>                    $c_msg = $xlsx_file_path + " ...Excelファイルオープン処理"
>                    cm_msg $c_msg
>                    ## Excelファイル整形：事前処理
>                    Add-Type -Path "${root_path}\module\EPPlus.dll"
>                    # ファイルコピー処理
>                    $src_xlsx = Import-Excel $xlsx_file_path
>                    # Excelファイル整形処理
>                    if($src_xlsx){
>                        ## Excelオブジェクト
>                        # Excelファイルオープン処理
>                        $excel = Import-Excel $xlsx_file_path
>                        # Excel Key/Valueオブジェクト追加
>                        $excel | Add-Member NoteProperty -Name "日付" -Value $YYMM
>
>                        # Excelファイル保存処理
>                        $excel | Export-Excel $dst_fs
>
>                        # ファイルチェック
>                        $ck_fs=  Compare-Object $xlsx_file_path $dst_fs -IncludeEqual
>                        cm_msg $ck_fs
>                    }
>                    else{
>                        $e_msg = "Excelファイル操作エラー"
>                        cm_msg $e_msg
>                    }
>                }
> ```

---

### 処理内容

#### 1.概要

本スクリプトは以下の処理ブロックにて処理を行います。

1. 事前処理
   1. ログ取得開始
   2. 外部ファンクション読み込み処理
2. ファイル整形処理
   1. コピー先フォルダ清掃
   2. CSV ファイル整形処理
   3. Excel ファイル整形処理
3. 事後処理
   1. 生成ファイル確認
   2. ログ出力停止

#### 1.事前処理

スクリプト実行に必要となる外部ファンクションの読み込みおよびログ取得の開始を行います。
ログ取得については Powershell 標準コマンドレット「Start-Transcript」にて実施します。
外部ファンクションは本スクリプトに同梱されている「functions」内より読み込み、外部モジュールとして「module」フォルダ内の"ImportExcel.psm1"よりコマンドレットを読み込みます。

#### 2.ファイル整形処理

事前処理の後に実行されます。
実行時、整形後ファイル格納フォルダの清掃処理(前ファイル削除処理)を行い、整形処理へ遷移します。
背系処理では CSV、Excel ファイル拡張子を判別し処理を分岐します。
今回のスクリプトでは対象ファイル指定があることからファイル名による判別も行っています。

Excel ファイル整形は外部ライブラリ「EPPlus.dll」を利用した処理を提供します。
その為、処理開始時に Microsoft .Net Core クラス定義にて該当 Dynamic Link Library の参照を行っています。

#### 3.事後処理

生成ファイルの確認およびログ取得の停止を行います。
ログ取得停止については Powershell 標準コマンドレット「Stop-Transcript」にて実施します。

本処理工程以外でエラーもしくは途中停止した場合、トランスクリプトが残ります。
その際は、手動で「Stop-Transcript」を実行して下さい。
