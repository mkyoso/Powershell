function cm_disconnect{
    param(
        [Parameter(ValueFromPipeLine =$true)]
        [ref]$Obj
    )
    $title = "Outlook�ؒf����"
    cm_header $title
    if ($Obj.Value -is [System.__ComObject]){
        $logger.info("Outlook�v���Z�X�ؒf")
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Obj.Value)|Out-Null
        $Obj = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()

        cm_footer
    }
}