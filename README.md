# AeroXL to PDF

**A beautiful batch Excel-to-PDF converter for Windows.**

AeroXL to PDF turns folders full of Excel workbooks into clean PDFs in a few clicks. It keeps your source-folder structure, shows live progress, and uses locally installed Microsoft Excel for faithful rendering.

![Platform](https://img.shields.io/badge/platform-Windows-19a9d6)
![Python](https://img.shields.io/badge/build-Python%203.8-20b37d)
![License](https://img.shields.io/badge/license-MIT-16b987)

## Features

- Batch conversion for `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`, and `.xltm`
- Recursive folder scan: process nested folders in one run
- Preserves the original folder structure in the PDF output directory
- Live progress, status, and a conversion activity log
- No administrator rights required
- Windows 8/8.1-friendly 32-bit release build
- Frutiger Aero-inspired desktop interface

## Requirements

- Windows 8, 8.1, 10, or 11
- Microsoft Excel installed locally. The application uses Excel's native PDF exporter, so complex formatting is rendered accurately.

## Run the app

1. Download [AeroXL2PDF_Win8.exe](https://github.com/JekoGeek/aeroxl-to-pdf/releases/download/v1.0.0/AeroXL2PDF_Win8.exe) from the Releases page.
2. Open the app.
3. Select the folder containing Excel files.
4. Select an output folder for PDFs.
5. Choose **Convert to PDF**.

[Download the latest release](https://github.com/JekoGeek/aeroxl-to-pdf/releases/latest)

## Build from source

The legacy-compatible build uses Python 3.8 32-bit:

```bat
py -3.8-32 -m pip install -r requirements.txt
build_win8.bat
```

The resulting executable is placed in `dist\AeroXL2PDF_Win8.exe`.

## Creator

Made by **JEKO**. Visit [My GitHub](https://github.com/JekoGeek).

## License

MIT. See [LICENSE](LICENSE).
