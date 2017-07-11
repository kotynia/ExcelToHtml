using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToHtml.CL
{
    class Program
    {
        static void Main(string[] args)
        {
            string ExcelPath = String.Join("", args);

#if DEBUG //TEST  
            ExcelPath = @"c:\git\ExcelToHtml\Test\test1.xlsx";
#endif

            string DataPath = ExcelPath + ".yaml";
            string HtmlPath = ExcelPath + ".html";

            Console.WriteLine("ExcelToHtml  https://github.com/marcinKotynia/ExcelToHtml ");
            Console.WriteLine("Usage: ExcelToHtml.exe [Path] ");
            Console.WriteLine("Usage: ExcelToHtml.exe c:\\book.xls ");



            try
            {
                //Read Excel File 
                FileInfo ExcelFile = new FileInfo(ExcelPath);


                var WorksheetHtml = new ExcelToHtml.ToHtml(ExcelFile);


                //Read Data Simple JSON cell,value
                FileInfo DataFile = new FileInfo(DataPath);
                if (!DataFile.Exists)
                {
                    Console.WriteLine("Data File Not Found {0}", DataFile);
                }
                else
                {

                    //Dictionary<string, string> Cells = new Dictionary<string, string>();
                    //InputOutput.Add("A1", "Hello World");  //set hello world
                    //InputOutput.Add("A2", "=2+1");  //set formula
                    //InputOutput.Add("[[TemplateField]]", "HelloTemplate");  //FillTempalte Filed
                    //InputOutput.Add(".A2", null);  //Output value form A2

                    string Data = File.ReadAllText(DataPath);

                    //Read Data From Yaml
                    var DeSerializer = new YamlDotNet.Serialization.Deserializer();
                    Dictionary<string, string> Cells = DeSerializer.Deserialize<Dictionary<string, string>>(Data);

                    //Fill Cells
                    var output = WorksheetHtml.GetSetCells(Cells);

                    var Serializer = new YamlDotNet.Serialization.Serializer();
                    string Yaml = Serializer.Serialize(output);

                    //For Debug Purpose output
                    File.WriteAllText(DataPath, Yaml);

                }


                string html = WorksheetHtml.Execute();

                File.WriteAllText(HtmlPath, html);

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR " + ex.Message);
                Console.ReadKey();
            }


        }
    }
}
