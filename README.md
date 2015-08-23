# unpackTunes.ps1
Windows PowerShell script to unpack randomly named music files and places them in folders by artist/album/title.
Script takes required input and output directory. For the input folder, it can have subfolders which the script will traverse.

## Sample input folder

:open_file_folder: c:\temp\tunes
<br/>&nbsp;&nbsp; :open_file_folder: F00
<br/>&nbsp;&nbsp;&nbsp;&nbsp; :camera: ADAN.m4a
<br/>&nbsp;&nbsp;&nbsp;&nbsp; :camera: BFVQ.mp3
<br/>&nbsp;&nbsp; :open_file_folder: F01
<br/>&nbsp;&nbsp;&nbsp;&nbsp; :camera: ALSR.mp3
<br/>&nbsp;&nbsp;&nbsp;&nbsp; :camera: PTZX.m4p

## Sample output folder
:open_file_folder: c:\music
<br/>&nbsp;&nbsp; :open_file_folder: Anoushka Shankar & Karsh Kale
<br/>&nbsp;&nbsp;&nbsp;&nbsp; :open_file_folder: Breathing Under Water
<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; :camera: Sea Dreamer (feat. Sting).mp4

<br/>&nbsp;&nbsp; :open_file_folder: Compilation
<br/>&nbsp;&nbsp;&nbsp;&nbsp; :open_file_folder: An Austin Rhythm And Blues Christmas
<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; :camera: Rockin' Around The Christmas Tree.m4a
<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; :camera: (Rockin') Winter Wonderland.m4a
<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; :camera: Boogie Woogie Santa Claus.m4a

## PowerShell command prompt examples
```powershell

> .\ unpackTunes.ps1 -musicSource C:\temp\tunes -musicDestination c:\music
#
# overwrite existing files
> .\ unpackTunes.ps1 -musicSource C:\temp\tunes -musicDestination c:\music -force
#
# remove the word "The" from artist name, such as "The Beatles" becomes "Beatles"
> .\ unpackTunes.ps1 -musicSource C:\temp\tunes -musicDestination c:\music -removeTheFromArtist -force

```

## Parameters
Parameter | Description
--------- | -------------
musicSource | **Required** folder or root folder for location of randomly named music files
musicDestination | **Required** target folder to copy files to
removeTheFromArtist | Remove the starting "The" from artist name, such as "The Beatles" becomes the "Beatles"
force | overwrite existing files


