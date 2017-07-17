using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelToHtml.console
{
    class Program
    {
        static void Main(string[] args)
        {

#if DEBUG  //TEST DATA 
            var testdata = new List<string>{
                @"-t=c:\git\ExcelToHtml\Test\test1.xlsx",
                @"-data=https://transit.land//api/v1/changesets/1/change_payloads"
            };
            args = testdata.ToArray();
#endif


            var arguments = ResolveArguments(args);
            string ExcelPath;
            string DataUrl;


            if (args.Length == 1)
                ExcelPath = args[0];
            else
                arguments.TryGetValue("-t", out ExcelPath);

            string DataPath = ExcelPath + ".yaml";
            string HtmlPath = ExcelPath + ".html";


            Console.WriteLine("ExcelToHtml https://github.com/marcinKotynia/ExcelToHtml");
            Console.WriteLine("ExcelToHtml.console.exe [xlsx File]");
            Console.WriteLine("");


#if !DEBUG
            try
            {
#endif

            Console.WriteLine(" Processing {0}", ExcelPath);

            //Read Excel File 
            FileInfo ExcelFile = new FileInfo(ExcelPath);


            var WorksheetHtml = new ExcelToHtml.ToHtml(ExcelFile);

            WorksheetHtml.DebugMode = true;

            //Read Data Simple JSON cell,value
            FileInfo DataFile = new FileInfo(DataPath);
            if (!DataFile.Exists)
            {
                Console.WriteLine(" Loading optional configuration {0} - Not found.", DataFile);
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

                //Get Set Cells and write to Yaml
                var output = WorksheetHtml.DataGetSet(Cells);
                var Serializer = new YamlDotNet.Serialization.Serializer();
                string Yaml = Serializer.Serialize(output);
                File.WriteAllText(DataPath, Yaml);
            }

            ExcelToHtml.console.Company test = new ExcelToHtml.console.Company();


            if (arguments.TryGetValue("-data", out DataUrl))
                WorksheetHtml.DataFromUrl(DataUrl);

            string html = WorksheetHtml.RenderHtml();

            Console.WriteLine(" File Saved {0}", HtmlPath);
            File.WriteAllText(HtmlPath, html);

#if DEBUG //DEBUG MODE  
            Console.ReadKey();
#endif

#if !DEBUG //DEBUG MODE
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR " + ex.Message);
                Console.ReadKey();
            }
#endif

        }


        /// <summary>
        /// Get Arguments
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static Dictionary<string, string> ResolveArguments(string[] args)
        {
            if (args == null)
                return null;

            if (args.Length > 0)
            {
                var arguments = new Dictionary<string, string>();

                for (int i = 0; i < args.Length; i++)
                {
                    int idx = args[i].IndexOf('=');
                    if (idx > 0)
                        arguments[args[i].Substring(0, idx)] = args[i].Substring(idx + 1);
                    else
                        arguments.Add(i.ToString(), args[i]);
                }

                return arguments;
            }

            return null;
        }

    }
}