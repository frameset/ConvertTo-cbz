<#
 .Synopsis
  Converts CBR files to CBZ files.

 .Description
  Converts CBR (comic book rar) files to CBZ (comic book zip) files from a given directory and subdirectories rescursively.
  Only works for Windows Powershell as it uses the UnRAR exe to decompress rar files.

 .Parameter Directory
  The directory that contains the cbr files you wish to convert.

 .Parameter UnRARLocation
  The location of the UnRAR.exe file. By default this will check the WinRAR package install directory.
  Alternatively you can download the UnRAR.exe from https://www.rarlab.com/rar/unrarw32.exe (this is a "self extracting archive" it runs as an executable that generates the actual UnRAR.exe)

 .Parameter DeleteOriginals
    If true, will delete the original cbr files.

 .Example
   # Convert a folder and subfolder of comics. This is a mandatory parameter.
   ConvertTo-Cbz -Directory "D:\Comics"

 .Example
   # Convert a folder and subfolder of comics, whilst specifying a particular UnRAR.exe
   ConvertTo-Cbz -Directory "D:\Comics" -UnRARLocation "D:\apps\unrar.exe"

 .Example
   # Convert a folder of comics, also delete the original cbr file.
   ConvertTo-Cbz -Directory "D:\Comics" -DeleteOriginals:$true

#>

param(
    [Parameter(Mandatory=$true)]
    [string] $directory
    [string] $unrarlocation = 'C:\Program Files\WinRAR\UnRAR.exe'
    [boolean] $deleteoriginals = $false
)

$startingdir = $directory
Set-Location $startingdir
$rars = Get-ChildItem -Recurse -Filter "*.cbr"
foreach($rar in $rars){
    $basepath = "$($rar.directoryname)\$($rar.basename)"
    if(-not(Test-Path $basepath)){mkdir $basepath}
    &"$unrarlocation" e -y $rar.fullname $basepath | Out-Null
    Get-ChildItem -LiteralPath $basepath | Compress-Archive -DestinationPath "$basepath.zip"
    Move-Item -LiteralPath "$basepath.zip" -Destination "$basepath.cbz"
    if($deleteoriginals) {Remove-Item -LiteralPath $basepath -Recurse -Force}
}