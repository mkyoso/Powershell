function cm_disconnect{
    param(
        [Parameter(ValueFromPipeLine =$true)]
        [ref]$Obj
    )
    $title = "Outlook切断処理"
    cm_header $title
    if ($Obj.Value -is [System.__ComObject]){
        $logger.info("Outlookプロセス切断")
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Obj.Value)|Out-Null
        $Obj = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()

        cm_footer
    }
}