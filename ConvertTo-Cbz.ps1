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
function ConvertTo-Cbz {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]
    [string]$directory,

    [string]$unrarlocation = 'C:\Program Files (x86)\WinRAR\UnRAR.exe',

    [boolean]$deleteoriginals = $false
  )
  
  $startingdir = Get-Location
  Set-Location $directory
  $rars = Get-ChildItem -Recurse -Filter "*.cbr"
  $i = 1
  foreach($rar in $rars){
    if ($rars -ne 1){
      Write-Progress -Activity "Processing cbr files" -Status "Converting $($rar.name)" -PercentComplete ($i/$rars.count*100)}
      $basepath = "$($rar.directoryname)\$($rar.basename)"
      if(-not(Test-Path $basepath)){mkdir $basepath}
      &"$unrarlocation" e -y $rar.fullname $basepath | Out-Null
      Get-ChildItem -LiteralPath $basepath | Compress-Archive -DestinationPath "$basepath.zip" |Out-Null
      Move-Item -LiteralPath "$basepath.zip" -Destination "$basepath.cbz" | Out-Null
      if($deleteoriginals) {Remove-Item -LiteralPath $basepath -Recurse -Force | Out-Null}
      #Start-Sleep -Seconds 2
      $i++
  }
  Set-Location $startingdir
}