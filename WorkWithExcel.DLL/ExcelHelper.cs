using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Diagnostics;

namespace WorkWithExcel.DLL
{
    public class ExcelHelper
    {
        private static Excel.Workbook excelBook = null;
        private static Excel.Application excelApp = null;
        private static Excel.Worksheet excelSheet = null;
        private static int convertedFiles = 0;

        public static int duplicateFiles = 0;    

                   
        private static void CSVToXlsx(string sourceFile, string xlsXFile)
        {
            //Read the file
            List<String> lines = new List<String>();
            using (StreamReader reader = new StreamReader(sourceFile, Encoding.GetEncoding("Windows-1251")))
            {
                string[] linesSplit = reader.ReadToEnd().Split(new[] { "\r\n" }, StringSplitOptions.None);

                lines.AddRange(linesSplit);               
            }

            //Now you got all lines of your CSV

            //Create your file with EPPLUS
            xlsXFile = GetNextFileName(xlsXFile);

            FileInfo xlsxFile = new FileInfo(xlsXFile);
            ExcelPackage pck = new ExcelPackage(xlsxFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets.Add("CSV");

            int i = 1;
            foreach (String line in lines)
            {
                int j = 1;
                var values = line.Split(';');
                foreach (String value in values)
                {
                    ws.Cells[i, j].Value = value;
                    j++;
                }
                i++;
            }

            pck.Save();
        }

        private static string GetNextFileName(string fileName)
        {
            string extension = Path.GetExtension(fileName);
            bool renamed = false;

            if (File.Exists(fileName))
            {
                renamed = true;
                duplicateFiles++;
            }                

            int i = 0;
          
            while (File.Exists(fileName))
            {
                Console.WriteLine("Файл с именем {0} уже существует", fileName);
                if (i == 0)
                    fileName = fileName.Replace(extension, "(" + ++i + ")" + extension);
                else
                    fileName = fileName.Replace("(" + i + ")" + extension, "(" + ++i + ")" + extension);
            }

            if (renamed)
            {
                Console.WriteLine("Файл был переименован в {0}", fileName);
            }
            
            return fileName;
        }

        public static void InitializeExcel()
        {
            excelApp = new Excel.Application { Visible = false };                                   
        }

        public static void CloseExcel()
        {
            excelApp.Quit();
        }

        public static void ConvertToXlsx(string sourceFile, string[] extensions, string destPath)
        {                            
            string fileExtension = Path.GetExtension(sourceFile);            

            foreach (string ext in extensions)
            {
                if (String.Equals(fileExtension, ext, StringComparison.CurrentCultureIgnoreCase))
                {
                    //string targetPath = destPath + "_" + fileExtension;
                    string targetPath = destPath;
                    Directory.CreateDirectory(targetPath);
                    string XLSXFile = Path.Combine(targetPath, Path.GetFileNameWithoutExtension(sourceFile) + ".xlsx");

                    if (String.Equals(fileExtension, ".csv", StringComparison.CurrentCultureIgnoreCase))
                    {
                        CSVToXlsx(sourceFile, XLSXFile);
                    }
                    else
                    {
                        try
                        {
                            excelBook = excelApp.Workbooks.Open(sourceFile, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);

                            XLSXFile = GetNextFileName(XLSXFile);

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
                    Console.WriteLine("Обработано файлов: {0}", ++convertedFiles);
                }
            }                                                 
        }
    }
}
