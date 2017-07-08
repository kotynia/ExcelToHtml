using System;
using System.IO;

namespace ExcelToHtml.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string localRepository = Directory.GetCurrentDirectory();

            Console.WriteLine("ExcelToHtml  https://github.com/marcinKotynia/ExcelToHtml ");
            Console.WriteLine("Usage: ExcelToHtml.exe [Path] ");
            Console.WriteLine("Usage: ExcelToHtml.exe c:\\book.xls ");


            string fullPath = String.Join("", args);

            try
            {
                Console.WriteLine(fullPath);
                FileInfo newFile = new FileInfo(fullPath);

                var WorksheetHtml = new ExcelToHtml.ConvertToHtml(newFile);
                string html = WorksheetHtml.ToHtml();

                File.WriteAllText(fullPath +".html", html);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR " + ex.Message);
                Console.ReadKey();
            }


        }
    }
}
