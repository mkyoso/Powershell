######################################################################
### .NetFrameworkAPI利用Office系アプリケーション操作支援PowerShell　###
######################################################################
Param(

)
######################################################################
### 関数定義
######################################################################

######################################################################
### 変数定義
######################################################################
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"

######################################################################
### 処理実行
######################################################################

### 事前処理
######################################################################
$rootPath = $PSScriptRoot

Write-Verbose ".NetFrameworkAPI利用Office系アプリケーション操作支援PowerShell Start"

#設定ファイル読み込み
$iniPath = "${rootPath}\settings\configure.ini"
Write-Host "--- iniファイル読み込み処理実施 ---"
if(Test-Path -PathType Leaf $iniPath){
    Write-Debug "設定ファイル読み込み $iniPath"
    $ini = @{}
    Get-Content $iniPath | %{ $ini += ConvertFrom-StringData $_ }
}
Write-Host "--- iniファイル読み込み処理終了 ---"
#Function読み込み
$read_func = Get-ChildItem "${rootPath}\functions" -Recurse -Filter *.ps1
Write-Host "--- 外部Function読み込み処理実施 ---"
foreach($func in $read_func){
    Write-Host "${func}ファイルを読み込みます"
    . $func.FullName
}
Write-Host "--- 外部Function読み込み処理終了 ---"

### 実行処理
######################################################################
### 1.logger処理
$log_base = $rootPath + "\log\"
$yyyymmdd = Get-Date -Format "yyyyMMdd"
$log_path = $log_base + "log_" + $yyyymmdd + ".log"
if (!(Test-Path -LiteralPath (Split-Path $log_path -parent) -PathType container)) {
    New-Item $log_path -type file -Force
}
#log4netのDLL読み込み
$dllFile = $rootPath + "\mod\log4net.dll";
Add-Type -Path $dllFile;
#log4net設定ファイル読み込み
$xmlFile = $rootPath + "\mod\log4net.xml";
$configFile = Get-Item $xmlFile;
[log4net.Config.XmlConfigurator]::Configure($configFile);
#ロガーの定義
$logger = [log4net.LogManager]::GetLogger($script:PSScriptRoot);
#ロガー設定
$rootLogger = ($logger.Logger.Repository).Root;
[log4net.Appender.FileAppender]$appender =  `
    [log4net.Appender.FileAppender]$rootLogger.GetAppender("FileAppender");
$appender.File = $log_path;
$appender.ActivateOptions();

### 2.主処理
$start_msg = "Outlookメールアイテム取得処理"
cm_header $start_msg

## Outlook接続処理
$ol = ol_connect
$dstpath = $rootPath + "\out"
## Outlookフォルダスキャン処理
$check_value = $ini["serch1"]
$check_value2 = $ini["serch2"]
if($check_value){
    $logger.info("${check_value}")
    if($check_value2){
        $logger.info("${check_value2}")
        ol_scanfolder $ol $dstpath $check_value $check_value2  -serch
    }
    else{
        ol_scanfolder $ol $dstpath $check_value -serch
    }
}
else{
    ol_scanfolder $ol $dstpath -all
}

### 事後処理
######################################################################
### 1.Outlookオブジェクト切断処理
[ref]$ol | cm_disconnect
cm_footer
