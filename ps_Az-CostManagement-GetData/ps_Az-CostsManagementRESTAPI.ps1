<#
Azure Cost Management REST API取得スクリプト
ファイル名：ps_Az-CostsManagementRESTAPI.ps1
作成日：2026/04/06
作成者：京増 誠志

.SYNOPSIS
Microsoft Azure Cost Management REST APIを実行し、結果レポートをCSV出力するスクリプト

.DESCRIPTION
本スクリプトは、Azureサブスクリプション内のリソースグループ毎にコスト情報を取得しAzure Portal
コスト管理画面にて確認できる情報をCSV整形してファイル出力するスクリプトとなります
本スクリプトにて利用しているライブラリは以下となります。

1.Azure PowerShell モジュール
Azure環境への接続・各種操作に使用するMicrosoft提供のコマンドレットモジュール
本スクリプトは本ライブラリが導入されていることが前提となります

# インストール方法
```
if (Get-Module -Name AzureRM -ListAvailable) {
    Write-Warning -Message ('Az module not installed. Having both the AzureRM and ' +
      'Az modules installed at the same time is not supported.')
} else {
    Install-Module -Name Az -AllowClobber -Scope CurrentUser
}
```

2. 実行時引数
.PARAMETER StartData
Azure コストマネジメント REST APIにて実コストを取得する指定期間の開始日(UTCタイムゾーン)を指定。
標準(引数指定なし)では7日前を設定しています。

.PARAMETER EndDate
Azure コストマネジメント REST APIにて実コストを取得する指定期間の終了日(UTCタイムゾーン)を指定。
標準(引数指定なし)では実行日の1日前を設定しています。

3. 実行例
.EXAMPLE
# 引数無し実行
. '\\{格納ディレクトリ}\AzGetCMData.ps1'

# 引数指定実行
.\AzGetCMData.ps1 -StartDate "yyyy-mm-dd" -EndDate "yyyy-mm-dd"

4. 参考リンク
.LINK
<https://qiita.com/tetsu_beer/items/c54a84f38c9aacf6c4ef>
#> 
 <# -- 変数定義 -- #>
param(     
    [Parameter(Mandatory = $false)]
    [string]$StartDate = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd"),
    [Parameter(Mandatory = $false)]
    [string]$EndDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
)

# 内部変数
## 変数クリア処理
$ResourceGroupName = ""

<# -- クラス定義
.SYNOPSIS
指定フォーマットにて実行スクリプトと同じディレクトリにログを生成するLoggerクラス

.DESCRIPTION
Loggerクラスは、各ログ種別でログを書き込むメソッドを提供

.PARAMETER encoding
ログファイルの文字エンコーディング
標準は"Default"
.PARAMETER logLayout
ログメッセージレイアウト
標準は"%d %p %m"
.PARAMETER datetimeLayout
日時フォーマット
標準は"yyyyMMddTHHmmssZ"

.LINK
<https://qiita.com/craftect/items/09637421f4df2c5eb57d>
-- #>
class Logger {
    ### properties
    [string] $encoding = "Default"
    [string] $logLayout = "%d %p %m"
    [string] $dataTimeLayout = "yyyyMMddTHHmmssZ"
    ### Log properties
    [string] $logCreatedDateTime
    [string] $logDir
    [string] $logFileName

    Logger ($InvokingScriptsMyInvocation){
        $this.logCreatedDateTime = Get-Date -Format yyyyMMdd
        $this.logDir = (Split-Path -Parent $InvokingScriptsMyInvocation.MyCommand.Path)
        $this.logFileName = (Split-Path -Leaf $InvokingScriptsMyInvocation.MyCommand.Name).Replace(".ps1","-" + $this.logCreatedDateTime + ".log")
    }

    [void] debug($msg){
        $this.writeLog($msg, "DEBUG")
    }
    [void] info($msg){
        $this.writeLog($msg, "INFO")
    }
    [void] warn($msg){
        $this.writeLog($msg, "WARNING")
    }
    [void] error($msg){
        $this.writeLog($msg, "ERROR")
    }
    [void] writeLog($msg,$logLevel){
        $replacedLayout = $this.logLayout
        $replacedLayout = $replacedLayout.Replace("%d", (Get-Date -Format $this.dataTimeLayout))
        $replacedLayout = $replacedLayout.Replace("%p", $logLevel)
        $replacedLayout = $replacedLayout.Replace("%m", $msg)

        Write-Output $replacedLayout | Out-File -FilePath (Join-Path $this.logDir $this.logFileName) -Append -Encoding $this.encoding
    }
}

<# -- 処理実行部 -- #>
# logger 起動
$logger = New-Object Logger($MyInvocation)

# モジュール読み込み
Import-Module Az.Accounts
Import-Module Az.Resources

# Azure認証確認・実行
## セッションキャッシュクリア処理
Clear-AzContext

## Azure認証処理
$context = Get-AzContext
if (-not $context) {
    $logger.Info("Azureにログインしています...")
    Connect-AzAccount
    $context = Get-AzContext
}

$logger.Info("コスト抽出対象サブスクリプション: $($context.Subscription.Name)")
$subscriptionId = $context.Subscription.Id

# 

# アクセストークン取得（Azure Management API用）
$logger.Info("Azure Management API認証トークンを取得しています...")
$token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate(
    $context.Account, 
    $context.Environment, 
    $context.Tenant.Id, 
    $null, 
    "Never", 
    $null, 
    "https://management.azure.com/"
).AccessToken

# リソースグループ抽出
$logger.Info("抽出対象リソースグループを取得します...")
try{
    $resourceGroups = Get-AzResourceGroup
    if($resourceGroups){
        foreach($rg in $resourceGroups.ResourceGroupName){
            # APIリクエスト設定（指定されたリソースグループと期間でコスト情報を取得）
            $logger.Info("期間: $StartDate から $EndDate")
            $logger.Info("対象リソースグループ: $($rg)")

            $ResourceGroupName = $rg
            $scope = "/subscriptions/$subscriptionId/resourceGroups/$ResourceGroupName"
            $requestBody = @{
                type = "ActualCost"                    # 実際のコスト情報
                timeframe = "Custom"                   # カスタム期間
                timePeriod = @{
                    from = "$StartDate" + "T00:00:00+00:00"    # 開始日時
                    to = "$EndDate" + "T23:59:59+00:00"        # 終了日時
                }
                dataset = @{
                    granularity = "Daily"              # 日次集計
                    aggregation = @{
                        totalCost = @{
                            name = "Cost"              # コスト合計
                            function = "Sum"
                        }
                    }
                    grouping = @(
                        @{ type = "Dimension"; name = "ResourceId" },     # リソース別
                        @{ type = "Dimension"; name = "ServiceName" }     # サービス別
                    )
                }
            }

            # API呼び出し実行
            $logger.Info("Cost Management APIを呼び出しています...")
            $requestBodyJson = $requestBody | ConvertTo-Json -Depth 10
            $uri = "https://management.azure.com$scope/providers/Microsoft.CostManagement/query?api-version=2021-10-01"
            $headers = @{ 
                'Authorization' = "Bearer $token"
                'Content-Type' = 'application/json' 
            }

            $response = Invoke-RestMethod -Uri $uri -Method Post -Body $requestBodyJson -Headers $headers

            # データ処理とCSV出力
            $logger.Info("取得したコストデータを処理しています...")
            $costData = @()
            foreach ($row in $response.properties.rows) {
                $costData += [PSCustomObject]@{
                    Cost = [decimal]$row[0]                                                          # コスト金額
                    UsageDate = ([DateTime]::ParseExact($row[1].ToString(), "yyyyMMdd", $null)).ToString('yyyy-MM-dd')  # 使用日
                    ResourceId = $row[2]                                                             # リソースID
                    ResourceName = ($row[2] -split '/')[-1]                                          # リソース名
                    ResourceType = if ($row[2] -match '/providers/([^/]+)/([^/]+)/') { "$($matches[1])/$($matches[2])" } else { "Unknown" }  # リソースタイプ
                    ServiceName = $row[3]                                                            # サービス名
                    Currency = $row[4]                                                               # 通貨
                }
            }

            # CSV出力（ファイル名にリソースグループ名を含む）
            $outputPath = "azure-CostsManagement_$ResourceGroupName.csv"
            $costData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
            $logger.Info("取得レコード数: $($costData.Count)")
            $logger.Info("コスト情報をCSVファイルに出力しました: $outputPath")
            Write-Host "コスト情報をCSVファイルに出力しました: $outputPath"
        }

    }
    else{
        $logger.warn("リソースグループが存在しません...")
        return
    }
}
catch{
    Write-Host "エラーが発生しました： $_"
    $logger.error("エラーが発生しました： $($_)")
    exit
}


