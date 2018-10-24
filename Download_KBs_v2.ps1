Import-Module .\Functions.ps1



$wantedkbs = @()
$WantedWin10KBs = @()
$KBsToDownload = @()
$KBsToDownload = @()

<#

    Run Script Below

#>


$FullDownload = $true
$DownloadOfficePatches = $true
$DownloadWindowsPatches = $true
$DownloadDotNetpatch = $false
$DownloadFlashPlayerPatch = $true

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
    
    $WantedOfficeKBs_Count = $WantedKBs.Count
}

#Or use a text file for the KBS we want
#$wantedkbs = Get-Content -Path "E:\Patches_PS_Script\03ARP18_ToDownload.txt

#Create this day's folder
$Date = Get-Date -Format "yyyMMdd"

if(-not (Test-Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\")){
    New-Item -ItemType Directory -Force -Name "$($Date)_OIF_WIMUpdates\" -Path "$($TempDriveRoot):\"
}

#region Copy all Previous KBs to new folder
if(Test-Path -Path $($LastUpdateDirectory)){
    Copy-Item -Path "$($LastUpdateDirectory)\*" -Destination "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\" -Recurse
}
#endregion

#region Make list of superseeded KBs that we can delete
    $KBsToDelete = ($outputObject | Where-Object {$_."Security KB superseded" -notlike "Not applicable"})."Security release KB article"
#endregion

#region Get list of all files currently in KB folder

    #Office KBs
    $AlreadyDownloadedKBs_Object = Get-ChildItem -Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\" -Filter "*.msp","*.cab","*.msu" -Recursive

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
    #foreach($KB in $KBsToDelete){
        #$AlreadyDownloadedKBs_Object | Where-Object { $_.BaseName -match $KB } | Remove-Item
    #}
#endregion

#region Remove the KBs that we already have from the ToDownload list
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

#region Add the flash player patch
    if($DownloadFlashPlayerPatch){
        #Hardcode the KB since flash does not update once a month....
        $KBsToDownload += "1709,4457146,Flash"

        #Currently 1607 isn't our target so don't bother downloading it
        #$KBsToDownload += "1607,4457146,Flash"


    }
#endregion




#Download the KBs that we need
$Count = 0

#Foreach Unique KB to download (No need to download the same KB twice...)
foreach($Patch in ($KBsToDownload | Select-Object -Unique)){
    if($Patch -like "*,*"){


        #foreach($Patch in $wantedkbs){
            $ProductFolder = $Patch.Split(",")[0]
            $KBNumber = $Patch.Split(",")[1]
            $ProductFamily = $Patch.Split(",")[2]
            if(-not (Test-Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\$($ProductFolder)")){
                New-Item -ItemType Directory -Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\" -Force -Name $ProductFolder
            }

            if($null -ne $ProductFamily){
                #This is a product specific to a Win10 version
                Invoke-DownloadLatestKB -KB "$($KBNumber) $($ProductFolder)" | Start-BitsTransfer -Destination "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\$($ProductFolder)\$($KBNumber)_$($ProductFolder)_$($ProductFamily).msu"
            }else{
                #This is just a Win10 CU
                Invoke-DownloadLatestKB -KB (Get-LatestKB -KB $KB).Product | Start-BitsTransfer -Destination "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\$($ProductFolder)\$($KB)_$($ProductFolder).msu"
                #Get-LatestKB -KB "$Patch" | Start-BitsTransfer -Destination "E:\$($Date)_OIF_WIMUpdates\$($ProductFolder)\$Patch.cab.zip"
            }
            
        #}
    }else{
        if(-not (Test-Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\Office2016_x86Updates")){
            New-Item -ItemType Directory -Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\" -Force -Name "Office2016_x86Updates"
        }

            Invoke-DownloadLatestKB -KB (Get-LatestKB -KB $Patch).KB | Start-BitsTransfer -Destination "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\Office2016_x86Updates\$Patch.cab"
    }

    Write-Progress -Activity "Downloading Patches" -Status "Complete:" -PercentComplete (($Count/$KBsToDownload.Count)*100)
    $Count += 1
}


#Verify Download
if($DownloadOfficePatches){
    $OfficePatches_Object = Get-ChildItem -Path "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates\Office2016_x86Updates" -Filter "*.cab"
    if($OfficePatches_Object.Count -eq $WantedOfficeKBs_Count){
        Write-Host "All Office Patches Accounted For. Starting Extraction"    
        #Extract Patches
        Extract-Updates -BasePath "$($TempDriveRoot):\$($Date)_OIF_WIMUpdates"

        #Delete File after we are done
        $OfficePatches_Object | Remove-Item
    }
}