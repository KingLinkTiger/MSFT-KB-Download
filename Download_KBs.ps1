Import-Module .\Extract-Updates.ps1

<#
.NOTES
    v1 Initial Version

    To Do
    Update building object to not utilize search and replace.
#>
function Get-LatestOfficeKB(){
    <#
    [string]$web = Invoke-WebRequest -Uri "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016" -usebasicparsing
    $web | Out-File E:\Patches_PS_Script\cache3.html
    $source = Get-Content -Path "E:\Patches_PS_Script\cache3.html" -Raw;
    #>


    $html = New-Object -ComObject "HTMLFile";
    [string]$source = Invoke-WebRequest -Uri "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016" -usebasicparsing
    $html.IHTMLDocument2_write($source);


    [regex]$regex = '<TABLE>[\s\S]*?<\/TABLE>'
    $tables = $regex.matches($html.body.innerHTML).groups.value

    ForEach($String in $tables[1]){
        $TableRows = $string -split '<tr.*?>'
        $TableRows = $TableRows | ForEach-Object{$_ -replace ","} | ForEach-Object{$_ -split "(?s)</T(?:D|H)>.*?<T(?:D|H).*?>" -join "," -replace "<(/?T(D|H|R|ABLE)|font).*?>" -replace "[\r\n]" -replace "</?EM>" -replace "</?SUP>" -replace "<BR>" -replace "</?STRONG>" -replace "</?TBODY>" -replace "</?A.*?>"} | ForEach-Object{$_.Trim(' ,')} | Where-Object {![string]::IsNullOrWhiteSpace($_)}
        $TableHeadder = $TableRows[0].Split(",")
    }


    #Build the output
    $outputObject = @()

    for($i=0;$i -lt $TableRows.Length;$i++){
        if($i -ne 0){
            $rowOut = New-Object -TypeName psobject 
            for($j=0;$j -lt $TableRows[$i].split(",").Length;$j++){
                $rowOut | Add-Member -MemberType NoteProperty -Name $TableHeadder[$j] -Value $TableRows[$i].split(",")[$j].trim()
            }
            $outputObject += $rowOut
        }
    }


 return $outputObject
}

function Get-InfoByUpdateID(){
    param($UpdateID = "")
    $output = ""

    [string]$web = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$($UpdateID)" -usebasicparsing


    $Replaced = ((($web -split "<div id=`"supersededbyInfo`" TABINDEX=`"1`" >")[1] -split "</div>")[0]).Trim()
    #Write-Host $Replaced
    $ReplacedPatchs = ((($web -split "<div id=`"supersededbyInfo`" TABINDEX=`"1`" >")[1] -split "<div style=`"padding-bottom: 0.3em;`">")[1] -split "</div>")[0]
    #Write-Host $ReplacedPatchs

    if($Replaced -eq "n/a"){
        #Write-Host "Latest Patch"
        #Write-Host $Product
        #Write-Host $UpdateID
        $output = New-Object PSObject -Property @{
           Product          = $Product
           UpdateID          = $UpdateID
           KB      = ($Product | Select-String -Pattern "KB[0-9]{7}").Matches[0]
        }
    }else{
        $Patch = $ReplacedPatchs -split "<a href='"

        #Get the last patch as this should be the latest one
        $Product = (([String]$Patch[-1].Trim() -split "'>")[1] -split "</a>")[0]
        $UpdateID = (([String]$Patch[-1].Trim() -split "updateid=")[1] -split "'>")[0]

        #Write-Host $Product
        #Write-Host $UpdateID
        Get-InfoByUpdateID -UpdateID $UpdateID
          
    }

    return $output
}
<#
.NOTES
    v1 Initial
    v2 Replaced split logic with built in HTML parsing. Create custom object to hold the Title and KBID at the same time and then return it if it is wanted.

    To Do
    Update Wanted logic to just specific Windows 10 Type and Architecture Type 
#>
function Invoke-DownloadLatestKB(){
    param($KB = "")
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

        $WantedPatch = (($null -ne ($WantedOfficeProducts | WHere-Object { $Product -match $_ })) -and ($Product -match "32-Bit Edition")) -or ($WantedProducts -contains $Product) -or ($null -ne ($WantedProducts | Where-Object { $Product -match $_ }))
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
}#End of Invoke-DownloadLatestKB

function Get-LatestKB(){
    param($KB = "")
    [string]$web = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$($KB)" -usebasicparsing
    #[string]$web = Get-Content -Path "E:\Patches_PS_Script\cache.xml"
    $output = ""

    $index = 0

    ((($web -split "<div id=`"tableContainer`" class=`"resultsBackGround`">")[1] -split "</table>")[0] -split "<tr") | ForEach-Object {
        if($index -eq 0 -or $index -eq 1){}else{
            $Product = (($_ -split ";`'>")[1] -split "</a>")[0].Trim()

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
                "Windows Defender on Windows 10 Version 1709 for x64-based Systems",
                "Security Update for Microsoft Office Excel Viewer 2007",
                "Microsoft Edge on Windows 10 Version 1803 for x64-based Systems",
                "Cumulative Update for Windows 10 Version 1803 for x64-based Systems",
                "Internet Explorer 11 on Windows 10 Version 1803 for x64-based Systems",
                "Adobe Flash Player on Windows 10 Version 1803 for x64-based Systems",
                "Windows Defender on Windows 10 Version 1803 for x64-based Systems"
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

            $WantedPatch = (($null -ne ($WantedOfficeProducts | Where-Object { $Product -match $_ })) -and ($Product -match "32-Bit Edition")) -or ($WantedProducts -contains $Product) -or ($null -ne ($WantedProducts | Where-Object { $Product -match $_ }))
            #Write-Host $Product
            #Write-Host $WantedPatch
            if($WantedPatch){
            
                $UpdateID = ((($_ -split "<input id=`"")[1] -split "`" class=`"flatLightBlueButton`" type=`"button`" value=`'Download`' />")[0]).Trim()
                $return = Get-InfoByUpdateID -UpdateID $UpdateID

                $output = New-Object PSObject -Property @{
                   Product          = $return.Product
                   UpdateID          = $return.UpdateID
                   KB      = $return.KB
                }
                return $output

                #Write-Host $Product
                #Write-Host $UpdateID
                #$PatchWeb = Invoke-WebRequest -Uri https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=fc4cdfde-7168-4a2e-b2df-0eff7fd27d6b
                #Write-Host $PatchWeb

            }else{
                Write-Output "Patch is not wanted"
            }
        }
        $index++
    }
}


$wantedkbs = @()
$WantedWin10KBs = @()
$KBsToDownload = @()

<#

    Run Script Below

#>


$FullDownload = $true
$DownloadOfficePatches = $true
$DownloadWindowsPatches = $false

$LastUpdateDirectory = "D:\20180821_OIF_WIMUpdates"
$TempDriveRoot = "D"


#Get the latest Windows 10 KBs
if($DownloadWindowsPatches){
    $WantedWin10KBs = Get-LatestKB -KB "$(Get-Date -Format "yyy-MM") Cumulative Update for Windows 10 Version for x64-based Systems"
}

if($DownloadOfficePatches){
    #Get the Office KBs
    $outputObject = Get-LatestOfficeKB

    #Remove the Office Language because we don't need it ever
    $outputObject = ($outputObject | Where-Object {$_.Product -ne "Office 2016 Language Interface Pack"})

    #Security KBs
    $wantedkbs = ($outputObject | Where-Object {$_."Security release KB article" -notlike "Not applicable"})."Security release KB article"

    #Non-Security KBs
    $wantedkbs += ($outputObject | Where-Object {$_."Non-security KB article" -notlike "Not applicable"})."Non-security KB article"
}

#Or use a text file for the KBS we want
#$wantedkbs = Get-Content -Path "E:\Patches_PS_Script\03ARP18_ToDownload.txt

#Create this day's folder
$Date = Get-Date -Format "yyyMMdd"

if(-not (Test-Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\")){
    New-Item -ItemType Directory -Force -Name "$($Date)_OIF_WIMUpdates\" -Path "$($TempDriveRoot):"
}

if(-not $FullDownload){
    #region Copy all Previous KBs to new folder
    Copy-Item -Path "$($LastUpdateDirectory)\*" -Destination "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\" -Recurse
    #endregion

    #region Make list of superseeded KBs that we can delete
    $KBsToDelete = ($outputObject | Where-Object {$_."Security KB superseded" -notlike "Not applicable"})."Security release KB article"
    #endregion

    #region Get list of all files currently in KB folder
    $AlreadyDownloadedKBs_Object = Get-ChildItem -Path "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\Office2016_x86Updates\"

    $AlreadyDownloadedKBs = @()
    ($AlreadyDownloadedKBs_Object).Name | ForEach-Object{
        $kb = ($_.split(".")[0])

        if($kb -match "_"){
            $kb = $kb.split("_")[0]
        
        }

        if($kb -match "KB"){
            $kb = $kb.replace("KB","")
        }

        $AlreadyDownloadedKBs += $kb
    }
    #endregion

    #region Delete the superseeded KBs as they're no longer needed.
    foreach($KB in $KBsToDelete){
        $AlreadyDownloadedKBs_Object | Where-Object { $_.BaseName -match $KB } | Remove-Item
    }
    #endregion

    #region Remove the KBs that we already have from the ToDownload list
    $KBsToDownload = @()
    foreach($KB in $wantedkbs){
        if($AlreadyDownloadedKBs -match $KB){
            #Do Nothing
        }else{
            #We want to download it.
            $KBsToDownload += $KB
        }
    }
    #endregion

    #region Add the Win10 Updates to the ToDownload List
    if($DownloadWindowsPatches){
        foreach($KB in $WantedWin10KBs){
            #$KB.Product
            $pattern = "Cumulative Update for Windows 10 Version (.*) for x64-based Systems"

            ($KB.Product -match $pattern).Groups
            $KBsToDownload += "$($Matches[1]),$($KB.KB)"
        }
    }
    #endregion
}else{
    #region Add the Win10 Updates to the ToDownload List
    if($DownloadWindowsPatches){
        foreach($KB in $WantedWin10KBs){
            #$KB.Product
            $pattern = "Cumulative Update for Windows 10 Version (.*) for x64-based Systems"

            ($KB.Product -match $pattern).Groups

            if($null -ne $($KB.KB)){
                $KBsToDownload += "$($Matches[1]),$($KB.KB)"
            }

        }
    }

    $KBsToDownload += $wantedkbs
}

#Download the KBs that we need
$Count = 0
foreach($Patch in $KBsToDownload){
    if($Patch -like "*,*"){


        #foreach($Patch in $wantedkbs){
            $ProductFolder = $Patch.Split(",")[0]
            $Patch = $Patch.Split(",")[1]
            if(-not (Test-Path "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\$($ProductFolder)")){
                New-Item -ItemType Directory -Path "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\" -Force -Name $ProductFolder
            }

            #Write-Host (Get-LatestKB -KB $Patch).KB
            Invoke-DownloadLatestKB -KB "$((Get-LatestKB -KB $Patch).KB) Cumulative Update for Windows 10 Version for x64-based Systems" | Start-BitsTransfer -Destination "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\$($ProductFolder)\$Patch.cab.zip"
            #Get-LatestKB -KB "$Patch" | Start-BitsTransfer -Destination "E:\$($Date)_OIF_WIMUpdates\$($ProductFolder)\$Patch.cab.zip"
        #}
    }else{
        if(-not (Test-Path "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\Office2016_x86Updates")){
            New-Item -ItemType Directory -Path "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\" -Force -Name "Office2016_x86Updates"
        }

            Invoke-DownloadLatestKB -KB (Get-LatestKB -KB $Patch).KB | Start-BitsTransfer -Destination "$($TempDriveRoot):$($Date)_OIF_WIMUpdates\Office2016_x86Updates\$Patch.cab.zip"
    }

    Write-Progress -Activity "Downloading Patches" -Status "Complete:" -PercentComplete (($Count/$KBsToDownload.Count)*100)
    $Count += 1
}

#Extract Patches
if($DownloadOfficePatches){
    Extract-Updates -BasePath "$($TempDriveRoot):$($Date)_OIF_WIMUpdates"
}