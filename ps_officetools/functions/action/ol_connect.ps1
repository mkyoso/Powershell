function ol_connect($proc_ck){
    try{
        $title = "Outlook接続処理"
        cm_header $title
        ## Processチェック
        $proc_ck = cm_checkprocess -outlook
        if($proc_ck -eq $True){
            $col = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
            $col.GetNameSpace("MAPI")
            if($col){
                $msg = "Connected successfully."
                $logger.info($msg)
                cm_footer
            }
        }
        else{
            Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
            $logger.info("Connecting to Outlook session")
            $col = New-object -ComObject Outlook.Application
            $col.GetNameSpace("MAPI")
            # MAPI Namespace https://msdn.microsoft.com/en-us/library/office/ff865800.aspx
            # Session https://msdn.microsoft.com/en-us/library/office/ff866436.aspx
            if($col){
                $msg = "Connected successfully."
                $logger.info($msg)
                cm_footer
            }
        }
    }
    catch{
        $logger.error("Can not obtain Outlook COM object. Try running Start-Outlook and then repeat command. "+($Error[0].Exception))
        throw($Error[0].Exception)
    }
}