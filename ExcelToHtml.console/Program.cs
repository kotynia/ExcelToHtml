using System;
using System.Collections.Generic;
using System.IO;
using ExcelToHtml.Helpers.WkHtmlToPdf;

namespace ExcelToHtml.console
{
    class Program
    {
        static void Main(string[] args)
        {

#if DEBUG  //TEST DATA 
            var testdata = new List<string>{
                @"-t=c:\git\ExcelToHtml\Test\test1.xlsx",
                @"-data=http://nflarrest.com/api/v1/crime"
               // @"-data=https://transit.land//api/v1/changesets/1/change_payloads",
               // @"-output=pdf"
            };
            args = testdata.ToArray();
#endif


            var arguments = ResolveArguments(args);
            string ExcelPath;
            string DataUrl;
            string Output = "html";
            string OutputPath;

            if (args.Length == 1)
                ExcelPath = args[0];
            else
                arguments.TryGetValue("-t", out ExcelPath);

            string DataPath = ExcelPath + ".yaml";


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

                if (!arguments.TryGetValue("-outputpath", out OutputPath))
                    OutputPath = ExcelPath;

                arguments.TryGetValue("-output", out Output);
                if (Output == null)
                    Output = "html";

                if (Output.ToLower() == "html")
                {
                    string html = WorksheetHtml.GetHtml();
                    File.WriteAllText(OutputPath + "." + Output, html);

                }
                else if (Output.ToLower() == "htmlw3css")
                {
                    string html = WorksheetHtml.GetHtml();
                    File.WriteAllText(OutputPath + ".html", String.Format(ExcelToHtml.Helpers.Strings.w3cssHTML, html));

                }
                else if (Output.ToLower() == "pdf")
                {
                    string html = String.Format(ExcelToHtml.Helpers.Strings.w3cssHTML, WorksheetHtml.GetHtml());
                    PdfConvert.ConvertHtmlToPdf(new PdfDocument
                    {
                        Html = html

                    }, new PdfOutput
                    {
                        OutputFilePath = OutputPath + "." + Output
                    });

                }
                else if (Output.ToLower() == "xlsx")
                {
                    File.WriteAllBytes(OutputPath + "." + Output, WorksheetHtml.GetBytes());
                }
                else
                {
                    throw new Exception("-output expected pdf,xlsx,html,htmlw3css");

                }

                Console.WriteLine(" File Saved {0}", OutputPath + "." + Output);


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