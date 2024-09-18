######################################################################
### .NetFrameworkAPI���pOffice�n�A�v���P�[�V��������x��PowerShell�@###
######################################################################
Param(

)
######################################################################
### �֐���`
######################################################################

######################################################################
### �ϐ���`
######################################################################
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"

######################################################################
### �������s
######################################################################

### ���O����
######################################################################
$rootPath = $PSScriptRoot

Write-Verbose ".NetFrameworkAPI���pOffice�n�A�v���P�[�V��������x��PowerShell Start"

#�ݒ�t�@�C���ǂݍ���
$iniPath = "${rootPath}\settings\configure.ini"
Write-Host "--- ini�t�@�C���ǂݍ��ݏ������{ ---"
if(Test-Path -PathType Leaf $iniPath){
    Write-Debug "�ݒ�t�@�C���ǂݍ��� $iniPath"
    $ini = @{}
    Get-Content $iniPath | %{ $ini += ConvertFrom-StringData $_ }
}
Write-Host "--- ini�t�@�C���ǂݍ��ݏ����I�� ---"
#Function�ǂݍ���
$read_func = Get-ChildItem "${rootPath}\functions" -Recurse -Filter *.ps1
Write-Host "--- �O��Function�ǂݍ��ݏ������{ ---"
foreach($func in $read_func){
    Write-Host "${func}�t�@�C����ǂݍ��݂܂�"
    . $func.FullName
}
Write-Host "--- �O��Function�ǂݍ��ݏ����I�� ---"

### ���s����
######################################################################
### 1.logger����
$log_base = $rootPath + "\log\"
$yyyymmdd = Get-Date -Format "yyyyMMdd"
$log_path = $log_base + "log_" + $yyyymmdd + ".log"
if (!(Test-Path -LiteralPath (Split-Path $log_path -parent) -PathType container)) {
    New-Item $log_path -type file -Force
}
#log4net��DLL�ǂݍ���
$dllFile = $rootPath + "\mod\log4net.dll";
Add-Type -Path $dllFile;
#log4net�ݒ�t�@�C���ǂݍ���
$xmlFile = $rootPath + "\mod\log4net.xml";
$configFile = Get-Item $xmlFile;
[log4net.Config.XmlConfigurator]::Configure($configFile);
#���K�[�̒�`
$logger = [log4net.LogManager]::GetLogger($script:PSScriptRoot);
#���K�[�ݒ�
$rootLogger = ($logger.Logger.Repository).Root;
[log4net.Appender.FileAppender]$appender =  `
    [log4net.Appender.FileAppender]$rootLogger.GetAppender("FileAppender");
$appender.File = $log_path;
$appender.ActivateOptions();

### 2.�又��
$start_msg = "Outlook���[���A�C�e���擾����"
cm_header $start_msg

## Outlook�ڑ�����
$ol = ol_connect
$dstpath = $rootPath + "\out"
## Outlook�t�H���_�X�L��������
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

### ���㏈��
######################################################################
### 1.Outlook�I�u�W�F�N�g�ؒf����
[ref]$ol | cm_disconnect
cm_footer
