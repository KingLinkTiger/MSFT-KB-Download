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
        $BasePath = "C:\Temp\20180816_OIF_WIMUpdates",
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