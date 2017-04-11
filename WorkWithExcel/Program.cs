using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using  WorkWithExcel.DLL;

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
            const string source = @"f:\Meta\";
            string[] masks = {".csv", ".ods", ".xls", ".xlsm" };

            try
            {
                if (Directory.Exists(source))
                {
                    ExcelHelper.InitializeExcel();
                    string[] metas = Directory.GetFiles(source, "*", SearchOption.AllDirectories);
                    int i = 1;                
                    foreach (string file in metas)
                    {
                        ExcelHelper.ConvertToXlsx(file, masks, Path.Combine(source, "xlsx"));
                        //Console.WriteLine("Обработано файлов: {0}", i);
                        i++;
                    }
                    Console.WriteLine("Дубликатов найдено: {0}", ExcelHelper.duplicateFiles);
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
            finally
            {
                ExcelHelper.CloseExcel();
            }          
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
