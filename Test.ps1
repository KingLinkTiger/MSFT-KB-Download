    $KB = "4461449"
    $kbObj = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$($KB)"

    #TableNumber is the index of the table on the page. Table 6 was found by trial and error
    $TableNumber = 6

    $tables = @($kbObj.ParsedHtml.getElementsByTagName("TABLE"))

    $table = $tables[$TableNumber]

    $ResultKBs = [System.Collections.ArrayList]::new()

    $i = 0
    foreach($row in @($table.Rows)){
        if($i -ne 0){
            $cells = @($row.Cells)
        
            $Title = $cells[1].InnerText.Trim()
            $KBID = $cells[1].id.Trim()
            $KBID = $KBID.Substring(0,$KBID.Length-6)

            $TempOutput = [pscustomobject]@{
                "Title" = $Title
                "KBID" = $KBID
            }
            [void]$ResultKBs.Add($TempOutput)
        }
        $i = $i+1
    }

    foreach($ResultKB in $ResultKBs){
        $WantedPatch = $false

        $WantedProducts = @(
            "Microsoft Edge on Windows 10 Version 1511 for x64-based Systems",
            "Cumulative Update for Windows 10 Version 1511 for x64-based Systems",
            "Internet Explorer 11 on Windows 10 Version 1511 for x64-based Systems",
            "Adobe Flash Player on Windows 10 Version 1511 for x64-based Systems",
            "Windows Defender on Windows 10 Version 1511 for x64-based Systems",
            "Microsoft Edge on Windows 10 Version 1607 for x64-based Systems",
            "Cumulative Update for Windows 10 Version 1607 for x64-based Systems",
            "Internet Explorer 11 on Windows 10 Version 1607 for x64-based Systems",
            "Adobe Flash Player on Windows 10 Version 1607 for x64-based Systems",
            "Windows Defender on Windows 10 Version 1607 for x64-based Systems",
            "Microsoft Edge on Windows 10 Version 1703 for x64-based Systems",
            "Cumulative Update for Windows 10 Version 1703 for x64-based Systems",
            "Internet Explorer 11 on Windows 10 Version 1703 for x64-based Systems",
            "Adobe Flash Player on Windows 10 Version 1703 for x64-based Systems",
            "Windows Defender on Windows 10 Version 1703 for x64-based Systems",
            "Microsoft Edge on Windows 10 Version 1709 for x64-based Systems",
            "Cumulative Update for Windows 10 Version 1709 for x64-based Systems",
            "Internet Explorer 11 on Windows 10 Version 1709 for x64-based Systems",
            "Adobe Flash Player on Windows 10 Version 1709 for x64-based Systems",
            "Windows Defender on Windows 10 Version 1709 for x64-based Systems"
            "Security Update for Microsoft Office Excel Viewer 2007"
        )

        $WantedOfficeProducts = @(
            "Microsoft Office 2016",
            "Microsoft Word 2016",
            "Microsoft Excel 2016",
            "Microsoft Outlook 2016",
            "Microsoft PowerPoint 2016",
            "Microsoft OneNote 2016",
            "Microsoft Visio 2016",
            "Microsoft Access 2016",
            "Microsoft Publisher 2016",
            "Microsoft Project 2016",
            "Microsoft OneDrive for Business",
            "Skype for Business 2016"
        )

        $WantedPatch = (($null -ne ($WantedOfficeProducts | ? { $Product -match $_ })) -and ($Product -match "32-Bit Edition")) -or ($WantedProducts -contains $Product) -or ($null -ne ($WantedProducts | ? { $Product -match $_ }))
        if($WantedPatch){
            $Post = @{ size = 0; updateID = $ResultKB.KBID; uidInfo = $ResultKB.KBID } | ConvertTo-Json -Compress
            $PostBody = @{ updateIDs = "[$Post]" } 

            Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -usebasicparsing -Method Post -Body $postBody |
                Select-Object -ExpandProperty Content |
                Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | 
                Select-Object -Unique |
                ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value }}  # Output for BITS
        }
    }#End of Foreach