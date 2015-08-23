
<#
.SYNOPSIS Splits out randomly named music files to artist/album/title
.DESCRIPTION Take input directory and output directory and process randomly named files to their artist/album/title
.PARAMETER musicSource = folder or root folder for location of randomly named music files
.PARAMETER musicDestination = folder to copy files to
.PARAMETER removeTheFromArtist = remove the starting "The" from artist name, such as "The Beatles" becomes the "Beatles"
.PARAMETER force = overwrite existing files
.EXAMPLE 
.\unpackTunes.ps1 -musicSource C:\temp\loonyTunes\music\ -musicDestination c:\data\music
.EXAMPLE 
Overwrite files with force flag   
.\unpackTunes.ps1 -musicSource C:\temp\loonyTunes\music\ -musicDestination c:\data\music -removeTheFromArtist -force
.LINK
.LINK
.NOTES
#>
param([string]$musicSource = $null, [string]$musicDestination = $null, [switch]$removeTheFromArtist, [switch] $force)



#from https://gallery.technet.microsoft.com/scriptcenter/c3d0ea6c-64a1-4716-a262-bcd71c9925fc
function funLine($strIN) {
    $strLine = "=" * $strIn.length
    Write-Host -ForegroundColor Yellow "`n$strIN"
    Write-Host -ForegroundColor Cyan $strLine
} #end funline


# from answer to http://stackoverflow.com/questions/23066783/how-to-strip-illegal-characters-before-trying-to-save-filenames
function remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)]
    [String]$Name
  )

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($Name -replace $re)
}


function get-artistName($metaData) {
    if (($metaData.ContainsKey("Part of a compilation")) -and ($metaData["Part of a compilation"] -eq "Yes")) {
        return "Compilation"
    }
    if ($metaData.ContainsKey("Album artist")) {
        return ($metaData["Album artist"])
    }
    if ($metaData.ContainsKey("Authors")) {
        return ($metaData["Authors"])
    }
    if ($metaData.ContainsKey("Contributing artists")) {
        return ($metaData["Contributing artists"])
    }
    if ($metaData.ContainsKey("Composers")) {
        return ($metaData["Composers"])
    }
    return "Unknown Artist"
}


function get-album($metaData) {
    if ($metaData.ContainsKey("Album")) {
        return ($metaData.Album)
    }
    return "Unknown Album"
}


# modified from https://gallery.technet.microsoft.com/scriptcenter/c3d0ea6c-64a1-4716-a262-bcd71c9925fc
function get-metadata($objFolder, $strFileName) {
    if ($strFileName.IsFolder -eq $true) {
        return $null
    }
    FunLine( "$($strFileName.Path)")
    $hash = @{ 'Path' = "$($strFileName.Path)" }
    for ($i = 0 ; $i  -le 266; $i++) { 
        if($objFolder.getDetailsOf($strFileName, $i)) {
            if($objFolder.getDetailsOf($objFolder.items, $i) -match "Composers|Authors|Contributing artists|Album|Title|Part of a compilation|File extension") {
                $hash.Add($($objFolder.getDetailsOf($objFolder.items, $i)), $($objFolder.getDetailsOf($strFileName, $i)))
            }
        }
    } 
    return $hash
}


function copy-musicFile($metaData, [string]$targetDir, [bool]$removeThe, [bool]$overwrite, $returnObj) {

    if ($metaData -eq $null) {
        return
    }
    $artist = get-artistName $metaData

    if ( ($removeThe -eq $true) -and ($artist.StartsWith("The ")) ) {
        $artist = $artist.Substring(4)
    }

    $artistPath = Join-Path $targetDir (remove-InvalidFileNameChars $artist)
    $album = get-album $metaData
    $albumPath = Join-Path $artistPath (remove-InvalidFileNameChars $album)
    
    if (-not (Test-Path -LiteralPath $albumPath)) {
        write-host $albumPath
        mkdir -Path $albumPath
    }

    if (-not ($metaData.ContainsKey("Title"))) {
        $metaData.add("Title", "Unknown")
    }
    $safeTitle = remove-InvalidFileNameChars ($metaData.Title)
    $titlePath = Join-Path $albumPath ($safeTitle + $metaData["File extension"])
    
    if ((-not (Test-Path -LiteralPath $titlePath)) -or ($overwrite -eq $true) ) {
        Copy-Item -Path ($metaData.Path) -Destination $titlePath -Force
        write-host "copied $titlePath"
        $returnObj.result++
    } else {     
        write-host "==========================================" -BackgroundColor DarkBlue -ForegroundColor Red
        write-host "already exists $titlePath " -BackgroundColor DarkBlue -ForegroundColor Red
        write-host "==========================================" -BackgroundColor DarkBlue -ForegroundColor Red
    }

}


function process-Folder([string]$sourceFolder, [string]$targetFolder, [bool]$removeThe, [bool]$overwrite, $returnObj) {
      
      $objShell = New-Object -ComObject Shell.Application
      $objFolder = $objShell.namespace($sourceFolder)
      
      foreach ($strFileName in $objFolder.items()) { 
            $fileMetaData = get-metadata $objFolder $strFileName
            copy-musicFile -metaData $fileMetaData -targetDir $targetFolder -removeThe $removeThe -overwrite $overwrite -returnObj $returnObj
      }
      return 
}


function unpackTunes-Main([string]$musicSource, [string]$musicTarget, [bool]$removeThe, [bool]$overwrite) {

    if (([string]::IsNullOrEmpty($musicSource)) -or ([string]::IsNullOrEmpty($musicTarget)) ) {
        Write-Error "You must specify musicSource and musicTarget"
        return
    }

    if (-not (Test-Path -LiteralPath $musicSource)) {
        Write-Error "Could not find $musicSource directory"
        return
    }
    
    if (-not (Test-Path -LiteralPath $musicTarget)) {
        mkdir $musicTarget
    }
    $fileCountObj = @{ result = 0 }
    
    #$totalCopied = 
    process-Folder -sourceFolder $musicSource -targetFolder $musicTarget -removeThe $removeThe -overwrite $overwrite -returnObj $fileCountObj

    $subfolders = gci -Path $musicSource -Directory -Recurse | Select-Object FullName
    foreach ($subfolder in $subfolders) {
        Write-Host ("processing " + ($subfolder.FullName) + " .....................") -BackgroundColor "DarkBlue" -ForegroundColor "Green"
        process-Folder -sourceFolder $subfolder.FullName -targetFolder $musicTarget -removeThe $removeThe -overwrite $overwrite -returnObj $fileCountObj
    }
    $msgCount = ("              Processed " + ($fileCountObj.result) + " files           ")
    write-host "/////////////////////////////////////////////////////" -BackgroundColor Green -ForegroundColor DarkBlue
    write-host $msgCount -BackgroundColor Green -ForegroundColor DarkBlue
    write-host "/////////////////////////////////////////////////////" -BackgroundColor Green -ForegroundColor DarkBlue

}



unpackTunes-Main -musicSource $musicSource -musicTarget $musicDestination -removeThe $removeTheFromArtist -overwrite $force
