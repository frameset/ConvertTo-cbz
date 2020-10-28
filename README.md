# ConvertTo-Cbz

ConvertTo-Cbz is a Windows Powershell script for converting comic book rar files (cbr) files to comic book zip (cbz) files.

## Installation

Download the .ps1 file from the respository.


## Usage

```powershell
   # Convert a folder and subfolder of comics. This is a mandatory parameter.
   ConvertTo-Cbz -Directory "D:\Comics"

   # Convert a folder and subfolder of comics, whilst specifying a particular UnRAR.exe
   ConvertTo-Cbz -Directory "D:\Comics" -UnRARLocation "D:\apps\unrar.exe"

   # Convert a folder of comics, also delete the original cbr file.
   ConvertTo-Cbz -Directory "D:\Comics" -DeleteOriginals:$true
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://choosealicense.com/licenses/mit/)