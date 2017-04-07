using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkWithExcel
{
    class Program
    {            
        static void Main(string[] args)
        {
            //command line parameters           
            //string arguments = String.Join(" ", args);   
            //string source = args[0];
            //string source = args[1];            
            string source = @"f:\Meta\Formats\";

            try
            {
                if (Directory.Exists(source))
                {
                    ExcelHelper.InitializeExcel();
                    string[] metas = Directory.GetFiles(source, "*", SearchOption.AllDirectories);
                    foreach (string file in metas)
                    {
                        ExcelHelper.ConvertToXlsx(file, Path.Combine(source, "xlsx"));
                    }
                    Console.WriteLine("Конвертация завершена");
                    Console.ReadLine();
                }                     
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null) ex = ex.InnerException;               
                Console.WriteLine("Произошла ошибка: {0}", ex.Message);
                Console.ReadLine();
            }


            //MyApp = new Excel.Application {Visible = false};
            //MyBook = MyApp.Workbooks.Open(DB_PATH);
            //MyBook.SaveAs("f:\\Historical_data_test\\AP Invoice 15 КП 2.4 (29.6.2016)\\2.4\\", Excel.XlFileFormat.xlExcel12,
            //System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
            //Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
            //System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            //excelWorkbook.SaveAs(strFullFilePathNoExt, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
            //Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
            //Excel.XlSaveConflictResolution.xlUserResolution, true,
            //Missing.Value, Missing.Value, Missing.Value);

            //MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            //lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

        }

        static void CopyFileByExtention(string filePath)
        {
            try
            {
                string rootPath = Directory.GetCurrentDirectory();
                string path = filePath;
                string ext = Path.GetExtension(path);
                string fileName = Path.GetFileName(path);

                if (File.Exists(path))
                {
                    if (!String.IsNullOrEmpty(ext))
                    {
                        Directory.CreateDirectory(Path.Combine(rootPath, ext));
                        File.Copy(path, Path.Combine(rootPath, ext, fileName));
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("The process failed: {0}", ex.ToString());
            }
        }
    }
}
