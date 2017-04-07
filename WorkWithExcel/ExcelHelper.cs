using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorkWithExcel
{
    class ExcelHelper
    {
        private static Excel.Workbook excelBook = null;
        private static Excel.Application excelApp = null;
        private static Excel.Worksheet excelSheet = null;

        public static void InitializeExcel()
        {
            excelApp = new Excel.Application { Visible = false };                                   
        }

        public static void ConvertToXlsx(string sourceFile, string destPath)
        {                            
            string fileExtension = Path.GetExtension(sourceFile);    

            if (fileExtension.Contains("csv") || fileExtension.Contains("ods") || fileExtension.Contains("xls") ||
                    fileExtension.Contains("xlsm"))
            {                    
                    string targetPath = destPath + "_" + fileExtension;
                    Directory.CreateDirectory(targetPath);
                    string XLSXFile = Path.Combine(targetPath, Path.GetFileNameWithoutExtension(sourceFile) + ".xlsx");
                    //excelApp = new Excel.Application {Visible = false};
                    excelBook = excelApp.Workbooks.Open(sourceFile, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                        Type.Missing, Type.Missing);
                    //excelSheet = (Excel.Worksheet)excelBook.Sheets[1]; // Explict cast is not required here    
                    try
                    {
                        excelBook.SaveAs(XLSXFile, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
                            Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                            Excel.XlSaveConflictResolution.xlUserResolution, true,
                            Missing.Value, Missing.Value, Missing.Value);
                    }
                    catch (Exception ex)
                    {
                        while (ex.InnerException != null) ex = ex.InnerException;
                        Console.WriteLine("Произошла ошибка: {0}", ex.Message);
                        Console.ReadLine();
                    }
                    finally
                    {
                        excelBook.Close();
                    }                                       
                }          
                    
        }
    }
}
