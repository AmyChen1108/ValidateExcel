using DocumentFormat.OpenXml.Packaging;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.IO;

namespace ConsoleApp_ValidateExcelFile
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            string filePath = "D:\\YourExcelFilePath\\excelfile.xlsx";
            ValidateExcelFile(filePath);
            ValidateExcelFileWithEPPlus(filePath);
            ValidateAndRepairExcelFile(filePath);
        }
        public static void ValidateExcelFile(string filePath)
        {
            // Console write filePath
            Console.WriteLine($"■ 檔案路徑: {filePath}");

            // Validate the file
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("File not found", filePath);
            }


            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Length == 0)
            {
                Console.WriteLine("檔案大小為 0，可能是空文件或上傳過程中出錯。");
                return;
            }

            string fileExtension = Path.GetExtension(filePath);
            // 驗證是否為 xlsx 或 xlsm 檔案
            if (fileExtension != ".xlsx" && fileExtension != ".xlsm")
            {
                throw new ArgumentException("Invalid file type. Only .xlsx and .xlsm files are supported", nameof(filePath));
            }
            //if (Path.GetExtension(filePath) != ".xlsx")
            //{
            //    throw new ArgumentException("Invalid file type. Only .xlsx files are supported", nameof(filePath));
            //}


            try
            {                
                // 讀取檔案並處理含有唯讀屬性的問題
                FileAttributes attributes = File.GetAttributes(filePath);

                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    Console.WriteLine("檔案是唯讀的，請確保具有正確的存取權限。");
                }

                // 如果是 .xlsm 檔案，提示使用者該檔案含有巨集
                if (fileExtension == ".xlsm")
                {
                    Console.WriteLine("檔案包含巨集，NPOI 不完全支援 .xlsm 格式。");
                    // 如果你只需讀取不涉及巨集的部分，仍可以嘗試讀取，但有可能會出現解析問題。
                }

                Console.WriteLine("========= Use NPOI 2.7.1 ========");

                // 嘗試讀取 Excel 檔案
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // 嘗試讀取 Excel 檔案
                    //XSSFWorkbook WK = new XSSFWorkbook(fs);
                    IWorkbook workbook = new XSSFWorkbook(fs);  // 讀取工作簿
                    Console.WriteLine("檔案讀取成功！");

                    // 進一步操作工作簿資料，例如遍歷工作表等
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        ISheet sheet = workbook.GetSheetAt(i);
                        Console.WriteLine($"工作表名稱: {sheet.SheetName}");
                    }
                }
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"讀取檔案時發生 IO 錯誤: {ioEx.Message}");
            }
            catch (InvalidOperationException invalidOpEx)
            {
                Console.WriteLine($"無效操作: {invalidOpEx.Message}");
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                Console.WriteLine($"無法開啟檔案，可能是 OLE 組合文檔: {comEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"無法讀取檔案: {ex.Message}");
            }


        }

        public static void ValidateExcelFileWithEPPlus(string filePath)
        {
            Console.WriteLine("");
            Console.WriteLine($"■ 檔案路徑: {filePath}");

            // 驗證檔案是否存在
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("File not found", filePath);
            }

            string fileExtension = Path.GetExtension(filePath).ToLower();

            // 驗證是否為 .xlsx 檔案
            if (fileExtension != ".xlsx")
            {
                throw new ArgumentException("Invalid file type. Only .xlsx files are supported", nameof(filePath));
            }

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                if (fileInfo.Length == 0)
                {
                    Console.WriteLine("檔案大小為 0，可能是空文件或上傳過程中出錯。");
                    return;
                }

                Console.WriteLine("========= Use EPPlus 7.4.0 ========");

                // 設定 LicenseContext 為 NonCommercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    if (workbook == null || workbook.Worksheets.Count == 0)
                    {
                        Console.WriteLine("檔案中沒有工作表。");
                        return;
                    }

                    Console.WriteLine("成功讀取 Excel 檔案！");
                    foreach (var sheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"工作表名稱: {sheet.Name}");
                    }
                }
            }  
            catch (IOException ioEx)
            {
                Console.WriteLine($"檔案讀取失敗: {ioEx.Message}");
            }
            catch (InvalidDataException dataEx)
            {
                Console.WriteLine($"資料格式錯誤: {dataEx.Message}");
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                Console.WriteLine($"無法開啟檔案，可能是 OLE 組合文檔: {comEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"無法讀取檔案: {ex.Message}");
            }
        }

        public static void ValidateAndRepairExcelFile(string filePath)
        {
            try
            {
                // 嘗試用 NPOI 打開 Excel 檔案
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook(fs);
                    Console.WriteLine("Excel 檔案讀取成功");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"無法讀取 Excel 檔案，錯誤訊息: {ex.Message}");

                // 如果遇到錯誤，嘗試用 OpenXML SDK 來檢查檔案
                if (ex.Message.Contains("EOF in header") || ex.Message.Contains("Wrong Local header signature"))
                {
                    Console.WriteLine("嘗試使用 OpenXML SDK 檢查檔案是否損壞...");
                    bool isValid = ValidateExcelFileWithOpenXML(filePath);
                    if (!isValid)
                    {
                        Console.WriteLine("檔案結構有問題或已損壞，請嘗試手動修復檔案或重新生成檔案。");
                    }
                    else
                    {
                        Console.WriteLine("檔案結構正常，請檢查其他問題。");
                    }
                }
            }
        }

        public static bool ValidateExcelFileWithOpenXML(string filePath)
        {
            try
            {
                // 打開檔案並檢查 OpenXML 檔案結構
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    // 檢查檔案是否符合 OpenXML 標準
                    var validationErrors = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheet>();
                    if (validationErrors == null)
                    {
                        Console.WriteLine("Excel 檔案結構無法檢測到工作表。");
                        return false;
                    }
                    Console.WriteLine("Excel 檔案結構檢測正常。");
                    return true;
                }
            }
            catch (OpenXmlPackageException oxEx)
            {
                Console.WriteLine($"OpenXML 檢查錯誤: {oxEx.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"無法檢查 Excel 檔案: {ex.Message}");
                return false;
            }
        }
    }
    
}
