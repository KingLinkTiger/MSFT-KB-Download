<#
.Synopsis
   Extract Microsoft Update CAB files into a single folder.
.DESCRIPTION
   Extract Microsoft Update CAB files into a single folder to easily update
   MDT images with updates.
.EXAMPLE
   Extract-Updates
.EXAMPLE
   Extract-Updates -Path "C:\Temp\20180816_OIF_WIMUpdates" -OfficePath "Office2016_x86Updates"
.INPUTS
   [String]Path
   [String]OfficePath
.OUTPUTS
   This function does not output anything.
.NOTES
    Old way to unzip the files using 7zip
    & "C:\Program Files\7-Zip\7zG.exe" x "$($BasePath)\$($FolderName)\$($_.BaseName).cab" -o"$($BasePath)\$($FolderName)\$($_.BaseName)"    
#>
function Extract-Updates{
    param(
        [Alias("UpdatesPath","Path")]                 
        $BasePath,
        [Alias("OfficeUpdatePath","OfficePath")] 
        $FolderName1 = "Office2016_x86Updates"
    )



    #Foreach .CAB.ZIP File rename to just .CAB
    $Count = 0
    $CABZIP_Object = Get-ChildItem -Path "$BasePath\$FolderName1" | Where-Object {$_.Name -like "*.cab.zip"}
    $CABZIP_Object | % {
        Move-Item -Path "$BasePath\$FolderName1\$_" -Destination "$($BasePath)\$($FolderName1)\$($_.BaseName)"
        $Count++
        Write-Progress -Activity "Renaming CAB Files" -Status "Renaming: $($_.BaseName)" -PercentComplete (($Count/$CABZIP_Object.Count)*100)
    }

    #Foreach CAB File
    $Count = 0
    $CAB_Object = Get-ChildItem -Path "$BasePath\$FolderName1" | Where-Object {$_.Name -like "*.cab"}
    $CAB_Object | % {
        if(-not (Test-Path -Path "$($BasePath)\$($FolderName1)\$($_.BaseName)")){
            New-Item -Path "$($BasePath)\$($FolderName1)\$($_.BaseName)" -ItemType Directory
        }

        Write-Host $_
        cmd.exe /c "C:\Windows\System32\expand.exe -F:*.msp $($BasePath)\$($FolderName1)\$($_.BaseName).cab $($BasePath)\$($FolderName1)\$($_.BaseName)"

        Write-Progress -Activity "Extracting CAB Files" -Status "Extracting: $($_.BaseName)" -PercentComplete (($Count/$CAB_Object.Count)*100)
        $Count++
    }


    #Create the All Updates folder if it does not exist
    if(-not (Test-Path -Path "$BasePath\$FolderName1\AllUpdates")){
        New-Item -Path "$BasePath\$FolderName1\AllUpdates" -ItemType Directory
    }

    #Foreach folder not called AllUpdates Get the MSP in it, rename it to the KB number, and move it to the AllUpdates Folder
    $Count = 0
    $Folder_Object = Get-ChildItem -Path "$BasePath\$FolderName1"  | Where-Object { $_.PSIsContainer -and $_.Name -ne "AllUpdates"}
    $Folder_Object | % {
        $KBNumber = $_.BaseName.split(".")[0]
        $FolderName = $_.BaseName

        Get-ChildItem -Path "$BasePath\$FolderName1\$($_.BaseName)" | Where-Object {$_.Name -like "*.msp*"} | % {
            #Rename the MSP to the KB number
            Rename-Item -Path "$BasePath\$FolderName1\$($FolderName)\$($_)" -NewName "$($KBNumber).msp"

            #Copy the MSP to the AllUpdates Folder
            Copy-Item -Path "$BasePath\$FolderName1\$($FolderName)\$($KBNumber).msp" -Destination "$BasePath\$FolderName1\AllUpdates\$($KBNumber).msp"

            #Remove the Folder the MSP was originally in
            Remove-Item -Path "$BasePath\$FolderName1\$($FolderName)\" -Recurse -Force
        }

        Write-Progress -Activity "Moving MSP Files" -Status "Moving: $($KBNumber)" -PercentComplete (($Count/$Folder_Object.Count)*100)
        $Count++
    }
}

<#
.NOTES
    v1 Initial Version - Moved to a function because of duplicate code

    To Do
#>
function Get-WantedPatch{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        $Title
    )

    $WantedPatch = $false

    $WantedWindows10Versions = @(
        #Currently 1607 is not our target so don't bother downloading it
        #"1607"
        "1709"
    )

    $WantedProducts = [System.Collections.ArrayList]::new()

    foreach($WantedWindows10Version in $WantedWindows10Versions){
        #[void]$WantedProducts.Add("Microsoft Edge on Windows 10 Version $($WantedWindows10Version) for x64-based Systems")
        [void]$WantedProducts.Add("Cumulative Update for Windows 10 Version $($WantedWindows10Version) for x64-based Systems")
        #[void]$WantedProducts.Add("Internet Explorer 11 on Windows 10 Version $($WantedWindows10Version) for x64-based Systems")
        #[void]$WantedProducts.Add("Adobe Flash Player on Windows 10 Version $($WantedWindows10Version) for x64-based Systems")
        #[void]$WantedProducts.Add("Windows Defender on Windows 10 Version $($WantedWindows10Version) for x64-based Systems")
        [void]$WantedProducts.Add("Security Update for Adobe Flash Player for Windows 10 Version $($WantedWindows10Version) for x64-based Systems")
    }

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

    foreach($WantedOfficeProduct in $WantedOfficeProducts){
        #Write-Host "$($ResultKB.Title) - $($WantedOfficeProduct)"
        if($Title -match $WantedOfficeProduct){
            if($Title -match "32-Bit Edition"){
                $WantedPatch = $true
            }
        }
    }

    foreach($WantedProduct in $WantedProducts){
        #Write-Host "$($ResultKB.Title) - $($WantedProduct)"
        if($Title -match $WantedProduct){
            $WantedPatch = $true
        }
    }

    return $WantedPatch
}#End Get-WantedPatch



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
v3 Updated to create a new HTML object instead of relying on INvoke-WebRequest to return the object... Because MSFT.
    Also completely fixed the broken WantedPatch code
    Removed WantedPatch code and moved to it's own function

To Do
Update Wanted logic to just specific Windows 10 Type and Architecture Type 
#>
function Invoke-DownloadLatestKB(){
    param($KB = "")

    $kbObj = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$($KB)" -UseBasicParsing

    # Create HTML file Object
    $HTML = New-Object -Com "HTMLFile"

    # Write HTML content according to DOM Level2 
    $HTML.IHTMLDocument2_write($kbObj.RawContent)

    #Replace kbObject with the HTML object
    $kbObj = $HTML


    #TableNumber is the index of the table on the page. Table 6 was found by trial and error
    $TableNumber = 6

    $tables = @($kbObj.getElementsByTagName("TABLE"))

    $table = $tables[$TableNumber]

    $ResultKBs = [System.Collections.ArrayList]::new()

    $i = 0
    foreach($row in @($table.Rows)){
        if($i -ne 0){
            $cells = @($row.Cells)
        
            $Title = $cells[1].InnerText.Trim()
            $KBID = $cells[1].id.Trim()
            
            [regex]$regex = "[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}"
            $KBID = $regex.Matches($KBID).Value

            #$KBID = $KBID.Substring(0,$KBID.Length-6)

            $TempOutput = [pscustomobject]@{
                "Title" = $Title
                "KBID" = $KBID
            }
            [void]$ResultKBs.Add($TempOutput)
        }
        $i = $i+1
    }

    foreach($ResultKB in $ResultKBs){
        $WantedPatch = (Get-WantedPatch -Title $ResultKB.Title)
        Write-Host $ResultKB.Title

        if($WantedPatch){
            $Post = @{ size = 0; updateID = $ResultKB.KBID; uidInfo = $ResultKB.KBID } | ConvertTo-Json -Compress
            $PostBody = @{ updateIDs = "[$Post]" } 

            Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -UseBasicParsing -Method Post -Body $postBody |
                Select-Object -ExpandProperty Content |
                Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | 
                Select-Object -Unique |
                ForEach-Object { 
                    [PSCustomObject] @{ 
                        Source = $_.matches.value 
                    }
                }  # Output for BITS
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

            $WantedPatch = (Get-WantedPatch -Title $Product)

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
                #Write-Output "Patch is not wanted"
            }
        }
        $index++
    }
}
