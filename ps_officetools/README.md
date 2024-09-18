# .NetFrameworkAPI利用Office系アプリケーション操作支援PowerShell
## Outlook受信メールフォルダ内アイテム取得用スクリプトサンプル

---
本スクリプトは.NetFrameworkモジュールを使用したOutlookメールアイテム取得スクリプトとなります。  
処理は.NetFramework OutlookAPIを利用します。  
> 現状はOutlookオブジェクトのみとなります。
  
実行時、Powershellセキュリティポリシー等にて実行できない場合は以下実行方式で使用が可能です。

> ` powershell
> powershell -ExecutionPolicy RemoteSigned ./ps_officetools.ps1
> `

### 利用方法
1)適当なディレクトリに本スクリプト一式を解凍します。  
　1-1)「settings」フォルダ内に格納されている"configure.ini"を開きます。 
　1-2)"configure.ini"内[serch1][serch2]に検索文字列を記述し上書き保存します。 
　※"configure.ini"内に検索文字列を記載しない場合は全件出力となります。 
2)「ps_officetools.ps1」を右クリックしメニューより"PowerShellで実行"。  
　2-1)権限により実行できない場合、以下コマンドプロンプト実行を実施。  
　※この際、別ユーザ管理のOutlookオブジェクト実行ではエラーとなりますので注意して下さい。  
> ` powershell
> powershell -ExecutionPolicy RemoteSigned ./ps_officetools.ps1
> `

3)本スクリプト内「out」フォルダ内に実行日付のフォルダが生成されていることを確認。  
4)実行日付フォルダ内に出力メールアイテムが存在していることを確認。  
