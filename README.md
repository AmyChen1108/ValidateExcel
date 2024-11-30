# ValidateExcel
![Issues](https://img.shields.io/github/issues/AmyChen1108/ValidateExcel)
![Language](https://img.shields.io/github/languages/top/AmyChen1108/ValidateExcel)
![License](https://img.shields.io/github/license/AmyChen1108/ValidateExcel)
![Updated](https://img.shields.io/badge/Last%20Updated-2024--11--30-red)
![](https://img.shields.io/badge/GitHub-AmyChen1108-blue?logo=github&style=social)

C#.net Console program. Check whether the excel document is a valid file.

# ConsoleApp_ValidateExcelFile

A .NET 8.0 console application for validating and repairing Excel files using **NPOI**, **EPPlus**, and **OpenXML SDK**.


## Features

- **Validate Excel files**: Ensures files are not corrupted, checks for valid formats (.xlsx, .xlsm), and detects issues like zero file size or read-only attributes.
- **Read Excel data**: Lists workbook sheets using **NPOI** and **EPPlus** libraries.
- **Repair damaged files**: Attempts to repair and validate Excel files using **OpenXML SDK** when possible.


## Technologies Used

- [.NET 8.0](https://dotnet.microsoft.com/)
- [DocumentFormat.OpenXml](https://github.com/OfficeDev/Open-XML-SDK) (v3.1.0)
- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) (v7.4.0)
- [NPOI](https://github.com/nissl-lab/npoi) (v2.7.1)


## Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) installed on your machine.
- An Excel file (e.g., `.xlsx` or `.xlsm`) to test the application.


## Getting Started

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/ConsoleApp_ValidateExcelFile.git
   cd ConsoleApp_ValidateExcelFile
2. Restore the required NuGet packages:
   ```bash
   dotnet restore
3. Run the application:
   ```bash
   dotnet run
4. Update the filePath variable in Program.cs to point to your local Excel file for testing.
 

## Example Usage
  ```bash
string filePath = "D:\\path\\to\\your\\file.xlsx";
ValidateExcelFile(filePath);          // Validate with NPOI
ValidateExcelFileWithEPPlus(filePath); // Validate with EPPlus
ValidateAndRepairExcelFile(filePath); // Attempt repair using OpenXML SDK
```

## Sample Console Output
  ```bash
■ 檔案路徑: D:\path\to\your\file.xlsx
========= Use NPOI 2.7.1 ========
檔案讀取成功！
工作表名稱: Sheet1
工作表名稱: Sheet2

========= Use EPPlus 7.4.0 ========
成功讀取 Excel 檔案！
工作表名稱: Sheet1
工作表名稱: Sheet2

檔案結構檢測正常。
  ```

## Error Handling

+ File not found: If the file path is invalid, the application will throw a FileNotFoundException.
+ Invalid file format: Only .xlsx and .xlsm files are supported.
+ Corrupted files: Attempts to repair with OpenXML SDK.


## License
This project uses EPPlus, which requires setting the LicenseContext to NonCommercial unless you have a commercial license. Ensure compliance with EPPlus licensing.
