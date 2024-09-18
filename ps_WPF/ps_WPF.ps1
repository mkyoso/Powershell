#########################################################
### Powershell管理用GUIツール
#########################################################
<#
    1.概要
    本スクリプトはPowershellの実行管理を一元化するためのGUIツールとなります。
    各画面でのアクション設定は本ファイル内の処理部を確認ください。

    Powershellの実行は本ツールフォルダ内、「functions」に格納されているスクリプト処理を想定しています。
    その為、実行ファイルはFunction化することが前提となっています。

    ■ 更新履歴
    ・202404dd ... 初版

    2.利用モジュール
    本スクリプトでは以下のアセンブリおよび外部モジュールを使用しています。

    ■ 利用アセンブリ
    ・PresentationCore
    ・PresentationFramework
    ・ReachFramework

    ■ 利用モジュール
    ・ImportExcel < https://github.com/dfinke/ImportExcel >

    3.スクリプト実行ができない場合
    3-1.実行時セキュリティエラーにて実行できない場合
    本スクリプト実行時、以下のセキュリティエラーが発生する場合は以下の処理を実行前に実施してください。

        1.エラー内容
        このシステムではスクリプトの実行が無効になっているため、ファイル ～.ps1 を読み込むことができません。
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
        > Set-ExecutionPolicy Bypass -Scope Process
        Bypass #実行権限の付与(Module読み込みのためRemotesignedよりも高い権限が必要)

    3-2.ネットワークドライブ(SPOドキュメント等)への参照が正常に行えない場合
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

        4.作成情報
        ・プログラム名：ps_wpf.ps1
        ・バージョン　：V1.0.0
        ・初回作成日　：2024/04/02
        ・最終更新日　：2024/MM/DD
        ・作成者　　　：Masashi Kyoso、masashi.kyoumasu@persol.co.jp

#>
#########################################################
# パラメーター定義
<#

#>
Param(
)

#########################################################
# 変数定義

$ErrorActionPreference = "stop"
Set-PSDebug -Strict

########################################
## 基本設定
$yyyymmdd = Get-Date -Format "yyyyMMdd"
$root_path = $PSScriptRoot
$Log_base = $root_path + "\Logs\"
$Log_path = $Log_base + "log_" + $yyyymmdd + ".log"
#$conf_base = $root_path + "\configs\"
#$conf_path = $conf_base + "config.ini"
$out_base = $root_path + "\out\"
$global:out_fs = $out_base + "diffs_" + $yyyymmdd + ".log"


#########################################################
# [事前処理]Log取得開始
Start-Transcript $Log_path -Append
Write-Verbose "スクリプトツール ... Start"

# Functions配下関数ファイル読み込み
$read_func = Get-ChildItem "${root_path}\functions" -Recurse -Filter *.ps1
Write-Host "外部関数呼び出し処理...Start"
foreach($func in $read_func){
    Write-Host "${func}ファイルを読み込みます"
    . $func.FullName
}

########################################
##アセンブリ読み込み処理
try{
    Add-Type -AssemblyName PresentationCore,PresentationFramework,ReachFramework,WindowsBase,System.Windows.Forms
}
catch{
    throw "Failed to load Windows Presentation Framework assemblies."
}

#########################################################
# [事前処理]GUI設定読み込み

########################################
## xamlファイル読込
<#
    \Views配下に設定しているXamlファイルの読み込み処理を行います。
    PowershellではC#等で読み取れるXAML属性も、Powershellでは読み取れないものがあるため読み取り時にリプレース処理を行っています。
#>
$PageFolder = $PSScriptRoot + "\Views\"
$main_path = $PageFolder + "Main.xaml"
$menu1_path = $PageFolder + "PreProcess.xaml"
$menu2_path = $PageFolder + "Process.xaml"
$menu3_path = $PageFolder + "PostProcess.xaml"
$inputXML = Get-Content $main_path -Raw
$input_form1 =  Get-Content $menu1_path -Raw
$input_form2 =  Get-Content $menu2_path -Raw
$input_form3 =  Get-Content $menu3_path -Raw
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
$input_form1 = $input_form1 -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
$input_form2 = $input_form2 -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
$input_form3 = $input_form3 -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[xml]$XAML = $inputXML
[xml]$form1 =  $input_form1
[xml]$form2 =  $input_form2
[xml]$form3 =  $input_form3

########################################
## XAMLデータ読込
<#
    XamlファイルデータをXAMLオブジェクトリーダーを使用して読み取りGUIを作成します。
#>
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
}
catch {
    Write-Warning $_.Exception
    throw
}
$PreProc_frame = (New-Object System.Xml.XmlNodeReader $form1)
try {
    $Pre_frame = [Windows.Markup.XamlReader]::Load( $PreProc_frame )
}
catch {
    Write-Warning $_.Exception
    throw
}
$Proc_frame = (New-Object System.Xml.XmlNodeReader $form2)
try {
    $Proc_frame = [Windows.Markup.XamlReader]::Load( $Proc_frame )
}
catch {
    Write-Warning $_.Exception
    throw
}
$PostProc_frame = (New-Object System.Xml.XmlNodeReader $form3)
try {
    $Post_frame = [Windows.Markup.XamlReader]::Load( $PostProc_frame )
}
catch {
    Write-Warning $_.Exception
    throw
}

########################################
## XAML属性の変数化
<#
    XAMLデータ内の属性を読み取りPowershell内で使用する変数として格納します。
#>
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)";
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
$form1.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)";
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $Pre_frame.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
$form2.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)";
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $Proc_frame.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
$form3.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)";
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $Post_frame.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}

#########################################################
# [処理]GUI処理設定

########################################
## Main_Page 設定
<#
    ■ 画面左部メニュー画面処理
    本処理はスタックパネル内ボタンクリックアクションを定義しています。
    ボタンアクションにより画面右部のフレーム内に各処理フォームを表示させています。
    フレーム内のGo/Back処理はフレームインスタンスに任せています。
#>
$var_pre_menu1.add_Click({
    $var_frame.Navigate($Pre_frame)
})
$var_proc_menu1.add_Click({
    $var_frame.Navigate($Proc_frame)
})
$var_post_menu1.add_Click({
    $var_frame.Navigate($Post_frame)
})

<#
    ■ 画面下部処理
    画面下部スクロールビューアーへのメッセージ出力処理を定義しています。
    各フレームページ処理より本処理を呼び出すことで出力を行います。
#>
$var_msg.Text = ""
Function println($line) {
    $var_msg.Inlines.Add($line + "`n")
    $var_scrollView.ScrollToBottom()
}
Function clearMsg() {
    $var_msg.Text = ""
}

########################################
## PreProcess.xaml_処理設定
<#
    ■ CSVファイル データグリッドビュー表示画面処理
    本処理はPreProcess.xaml内の処理を定義しています。
    ファイル選択ボタンのクリックアクションにより、選択ダイアログのファイル情報を取得し、実行処理となる"act_csv"へ情報引き渡しを行います。
#>

$var_ogv_fs_btn.add_Click({
    $ofd_fs = New-Object System.Windows.Forms.OpenFileDialog
    $ofd_fs.Filter = "CSV ファイル(*.CSV)|*.CSV"
    $ofd_fs.Title = "ファイルを選択してください"
    if($ofd_fs.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $global:ogv_fs = $ofd_fs.FileName
        println($ogv_fs + " が選択されました。")
    }
})

$var_ogv.add_Click({
    act_csv $ogv_fs
})

########################################
## Process.xaml_処理設定
<#
    ■ コンペア処理画面処理
    本処理はProcess.xaml内の処理を定義しています。
    ファイル選択ボタンのクリックアクションにより、選択ダイアログのファイル情報を取得し実行処理となる"act_diffs"へ情報引き渡しを行います。
    メッセージ出力については"Main.xaml"下部のスクロールビューワーへ出力表示させています。
#>

$var_src_fs_btn.add_Click({
    $src_fs = New-Object System.Windows.Forms.OpenFileDialog
    $src_fs.Title = "ファイルを選択してください"
    if($src_fs.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $global:fromfs = $src_fs.FileName
        println($fromfs + " が選択されました。")
    }
})
$var_dst_fs_btn.add_Click({
    $dst_fs = New-Object System.Windows.Forms.OpenFileDialog
    $dst_fs.Title = "ファイルを選択してください"
    if($dst_fs.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $global:tofs = $dst_fs.FileName
        println($tofs + " が選択されました。")

    }
})
$var_diffs.add_Click({
    $diffs = act_diffs $fromfs $tofs
    println($diffs)
})

########################################
## PostProcess.xaml_処理設定
<#
    ■ 実行端末情報表示画面処理
    本処理はPostProcess.xaml内の処理を定義しています。
    ここではOS情報を"Get-CimInstance"を経由し取得しています。
    取得結果は画面内のラベルに表示されます。
#>
$totalRam = [Math]::Round((Get-CimInstance -Class win32_computersystem).TotalPhysicalMemory/1Gb)
$totalCPU = (Get-CimInstance –ClassName Win32_Processor).NumberOfLogicalProcessors
$props = 'DriveLetter',
        @{Name = 'SizeRemainingGB'; Expression = {"{0:N2}" -f ($_.SizeRemaining/ 1Gb)}},
        @{Name = 'SizeGB'; Expression = {"{0:N2}" -f ($_.Size / 1Gb)}},
        @{Name = '% Free'; Expression = {"{0:P}" -f ($_.SizeRemaining / $_.Size)}}

$Diskmgmt = Get-Volume | Select-Object $props | Sort-Object DriveLetter | Format-Table
$var_lbl_Hostname.Content = $env:COMPUTERNAME
$var_lbl_Ram.Content = "$($totalRam) GB"
$var_lbl_CPUCores.Content = "$($totalCPU) Logical Cores"
$var_lbl_diskInfo.Text = $Diskmgmt | Out-String

$window.ShowDialog() | Out-Null