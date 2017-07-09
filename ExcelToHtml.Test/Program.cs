using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToHtmlConsole
{
    class Program
    {
        static void Main(string[] args)
        {


            //string localRepository = Directory.GetCurrentDirectory();
            //string fullPath = String.Join("", args); 

            //TEST!!!
            string ExcelPath = @"c:\git\ExcelToHtml\Test\Book1.xlsx";
            string DataPath = ExcelPath + ".yaml";
            string HtmlPath = ExcelPath + ".html";

            Console.WriteLine("ExcelToHtml  https://github.com/marcinKotynia/ExcelToHtml ");
            Console.WriteLine("Usage: ExcelToHtml.exe [Path] ");
            Console.WriteLine("Usage: ExcelToHtml.exe c:\\book.xls ");

            //try
            //{
                //Read Excel File 
                FileInfo ExcelFile = new FileInfo(ExcelPath);


                var WorksheetHtml = new ExcelToHtml(ExcelFile);


                //Read Data Simple JSON cell,value
                FileInfo DataFile = new FileInfo(DataPath);
                if (!DataFile.Exists)
                {
                    Console.WriteLine("Data File Not Found {0}", DataFile);
                }
                else
                {
                    //Read Data From Yaml
                    Dictionary<string, string> values = new Dictionary<string, string>();

                    string Data = File.ReadAllText(DataPath);

                    var DeSerializer = new YamlDotNet.Serialization.Deserializer();
                    Dictionary<string, string> Cells = DeSerializer.Deserialize<Dictionary<string, string>>(Data);

                    //Fill Cells
                    var output = WorksheetHtml.GetSetCells(Cells);

                    var Serializer = new YamlDotNet.Serialization.Serializer();
                    string Yaml = Serializer.Serialize(output);

                    //For Debug Purpose output
                    File.WriteAllText(DataPath,Yaml);

                }


                string html = WorksheetHtml.ToHtml();

                File.WriteAllText(HtmlPath, html);

            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("ERROR " + ex.Message);
            //    Console.ReadKey();
            //}


        }
    }
}
