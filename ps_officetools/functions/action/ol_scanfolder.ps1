function ol_scanfolder{
    Param(
        [Parameter()]
        $olObj,
        [Parameter()]
        $dstDir,
        [Parameter()]
        [string]$serch1,
        [Parameter()]
        [string]$serch2,
        [Parameter()]
        [Switch]
        $all,
        [Parameter()]
        [Switch]
        $serch
    )
    $title = "Outlook受信フォルダ検索処理"
    cm_header $title
    $logger.info("${serch1}")
    $logger.info("${serch2}")
    if($all){
        Add-type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
        $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
        $msg = "Outlook受信フォルダスキャン処理"
        $logger.info($msg)
        $olFolder = $olObj.getDefaultFolder($olFolders::olFolderInBox)
        $yymmdd = (Get-Date).ToString("yyyyMMdd")
        $dirname = $dstDir + "\" + $yymmdd
        if(Test-Path $dirname){
            $logger.info("${dirname} is found!")
        }
        else{
            $logger.info("${dirname} is not found!")
            New-Item $dirname -type Directory
        }
        $itemCollection = $olFolder.Items
        if($itemCollection.Count -eq 0){
            $logger.warn("件数：0")
        }
        else{
            $itemcont = $itemCollection.Count
            $logger.warn("件数：" + $itemcont)
        }
        $ckday = Get-Date
        $ckmon = $ckday.AddMonths(-1)
        $olitems = $itemCollection | Where-Object {$_.ReceivedTime -gt $ckmon}
        foreach($olItem in $olitems){
            $flname = $olItem.Subject
            $act_msg = "処理件名:" + $flname
            $logger.info($act_msg)
            cm_msg $act_msg
            $data = @{
                "subject" = $olitem.Subject;
                "body" = $olitem.Body;
                "htmlBody" = $olitem.HTMLBody;
            };
            $ml_subject = $olItem.Subject.Replace('/', '／').Replace('\', '￥'). `
            Replace('<', '＜').Replace('>', '＞').Replace('*', '＊').Replace('?', '？'). `
            Replace('|', '｜').Replace(':', '：').Replace(';', '；').Replace('[', '［'). `
            Replace(']', '］').Replace('"', '”')
            $message = $dirname + "\" + $ml_subject + "message.json"
            $data | ConvertTo-Json | Out-File -FilePath $message -Encoding utf8
        }
    }
    elseif($serch){
        if($serch2){
            Add-type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
            $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
            $msg = "Outlook受信フォルダスキャン処理"
            $logger.info($msg)
            $olFolder = $olObj.getDefaultFolder($olFolders::olFolderInBox)
            $yymmdd = (Get-Date).ToString("yyyyMMdd")
            $dirname = $dstDir + "\" + $yymmdd
            if(Test-Path $dirname){
                $logger.info("${dirname} is found!")
            }
            else{
                $logger.info("${dirname} is not found!")
                New-Item $dirname -type Directory
            }
            $itemCollection = $olFolder.Items
            if($itemCollection.Count -eq 0){
                $logger.warn("件数：0")
            }
            else{
                $itemcont = $itemCollection.Count
                $logger.warn("件数：" + $itemcont)
            }
            $ckday = Get-Date
            $ckmon = $ckday.AddMonths(-1)
            $olitems = $itemCollection | Where-Object {$_.ReceivedTime -gt $ckmon}
            foreach($olItem in $olitems){
                $check_str = "*" + $serch1 + "*"
                $check_str2 = "*" + $serch2 + "*"
                if($olItem.Subject -like $check_str -and $olItem.Subject -like $check_str2){
                    $flname = $olItem.Subject
                    $act_msg = "処理件名:" + $flname
                    $logger.info($act_msg)
                    cm_msg $act_msg
                    $data = @{
                        "subject" = $olitem.Subject;
                        "body" = $olitem.Body;
                        "htmlBody" = $olitem.HTMLBody;
                    };
                    $ml_subject = $olItem.Subject.Replace('/', '／').Replace('\', '￥'). `
                    Replace('<', '＜').Replace('>', '＞').Replace('*', '＊').Replace('?', '？'). `
                    Replace('|', '｜').Replace(':', '：').Replace(';', '；').Replace('[', '［'). `
                    Replace(']', '］').Replace('"', '”')
                    $message = $dirname + "\" + $ml_subject + "message.json"
                    $data | ConvertTo-Json | Out-File -FilePath $message -Encoding utf8
                }
            }
        }
        else{
            Add-type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
            $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
            $msg = "Outlook受信フォルダスキャン処理"
            $logger.info($msg)
            $olFolder = $olObj.getDefaultFolder($olFolders::olFolderInBox)
            $yymmdd = (Get-Date).ToString("yyyyMMdd")
            $dirname = $dstDir + "\" + $yymmdd
            if(Test-Path $dirname){
                $logger.info("${dirname} is found!")
            }
            else{
                $logger.info("${dirname} is not found!")
                New-Item $dirname -type Directory
            }
            $itemCollection = $olFolder.Items
            if($itemCollection.Count -eq 0){
                $logger.warn("件数：0")
            }
            else{
                $itemcont = $itemCollection.Count
                $logger.info("件数：" + $itemcont)
            }
            $ckday = Get-Date
            $ckmon = $ckday.AddMonths(-1)
            $olitems = $itemCollection | Where-Object {$_.ReceivedTime -gt $ckmon}
            foreach($olItem in $olitems){
                $check_str = "*" + $serch1 + "*"
                if($olItem.Subject -like $check_str ){
                    $flname = $olItem.Subject
                    $act_msg = "処理件名:" + $flname
                    $logger.info($act_msg)
                    cm_msg $act_msg
                    $data = @{
                        "subject" = $olitem.Subject;
                        "body" = $olitem.Body;
                        "htmlBody" = $olitem.HTMLBody;
                    };
                    $ml_subject = $olItem.Subject.Replace('/', '／').Replace('\', '￥'). `
                    Replace('<', '＜').Replace('>', '＞').Replace('*', '＊').Replace('?', '？'). `
                    Replace('|', '｜').Replace(':', '：').Replace(';', '；').Replace('[', '［'). `
                    Replace(']', '］').Replace('"', '”')
                    $message = $dirname + "\" + $ml_subject + "message.json"
                    $data | ConvertTo-Json | Out-File -FilePath $message -Encoding utf8
                }
            }
        }
    }
}