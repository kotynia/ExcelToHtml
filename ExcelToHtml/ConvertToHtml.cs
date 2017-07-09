using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

public class ExcelToHtml
{
    /// <summary>
    /// If not specified first used
    /// </summary>
    private string WorksheetName = String.Empty;
    ExcelPackage Excel;
    ExcelWorksheet WorkSheet;
    private Dictionary<string, string> TemplateFieldList;

    public string option_tableAtribute = "border-collapse: collapse;font-family: helvetica, arial, sans-serif;";

    public ExcelToHtml(FileInfo excelFile, string WorkSheetName = null)
    {
        if (!excelFile.Exists)
            throw new Exception(String.Format("File {0} Not Found", excelFile.FullName));

        Excel = new ExcelPackage(excelFile);
        if (!string.IsNullOrEmpty(WorkSheetName))
            WorkSheet = Excel.Workbook.Worksheets[WorksheetName];
        else
            WorkSheet = Excel.Workbook.Worksheets[1];
    }


    /// <summary>
    /// Render HTML
    /// </summary>
    /// <returns>Html</returns>
    public string ToHtml()
    {
        //GET TIME
        var watch = System.Diagnostics.Stopwatch.StartNew();

        //GET DIMENSIONS
        var start = WorkSheet.Dimension.Start;
        var end = WorkSheet.Dimension.End;
        StringBuilder sb = new StringBuilder();

        // Row by row
        for (int row = start.Row; row <= end.Row; row++)
        {
            sb.AppendLine("<tr>");
            for (int col = start.Column; col <= end.Column; col++)
            {
                var y = WorkSheet.Column(col);
                var d = WorkSheet.Cells[row, col];

                int merged = 0;

                if (d.Merge) //row is merged
                    merged = d.Worksheet.SelectedRange[WorkSheet.MergedCells[row, col]].Columns;

                //11 default font size
                var x = CellHelper.Process(WorkSheet.Cells[row, col], WorkSheet.Column(col).Width, 11, merged);
                sb.AppendLine(x);
                if (d.Merge)
                    col += (merged - 1);
            }
            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table>");
        watch.Stop();
        var elapsedMs = watch.ElapsedMilliseconds;

        return string.Format("<table  style=\"{0}>\" data-cth-ms=\"{1}\" data-cth-date=\"{2}\">{3}</table>",
            option_tableAtribute, elapsedMs, DateTime.Now, sb.ToString());
    }



    /// <summary>
    /// Set Cell Values
    /// Supported Format 
    /// A4 Test Value
    /// A5 =A2+A3
    /// [[TempalteField]] Test template
    /// </summary>
    /// <param name="data"></param>
    public Dictionary<string, string> GetSetCells(Dictionary<string, string> data)
    {
        //Dicionary to Excel
        foreach (var item in data)
        {
            //Template Handler
            if (item.Key.StartsWith("."))
            {
                //output handler usefull for tests
            }
            else if (item.Key.StartsWith("[["))
            {

                CreateTemplateFieldMap(); //One time

                var FieldList = TemplateFieldList.Where(x => x.Value == item.Key);

                foreach (var cellTemplate in FieldList)
                {
                    if (item.Value.StartsWith("=")) //Formula 
                        WorkSheet.Cells[cellTemplate.Key].Formula = item.Value.Remove(0, 1);
                    else //Text
                        WorkSheet.Cells[cellTemplate.Key].Value = item.Value;
                }


            }
            else if (item.Value.StartsWith("=")) //Formula 
                WorkSheet.Cells[item.Key].Formula = item.Value.Remove(0, 1);
            else //Text
                WorkSheet.Cells[item.Key].Value = item.Value;

        }

        //Probably Important to Calculate before get 
        WorkSheet.Calculate();

        string[] keys = data.Keys.ToArray();



        //Fill Out return values
        foreach (var item in keys)
        {
            if (item.StartsWith("."))
            {
                data[item] = WorkSheet.Cells[item.Remove(0, 1)].Text;
            }

        }

        return data;

    }
    

    /// <summary>
    /// Create Template Field Map, sample: [[Test]]
    /// </summary>
    public void CreateTemplateFieldMap()
    {

        if (TemplateFieldList == null)
        {
            TemplateFieldList = new Dictionary<string, string>();

            var start = WorkSheet.Dimension.Start;
            var end = WorkSheet.Dimension.End;

            for (int row = start.Row; row <= end.Row; row++)
            {
                for (int col = start.Column; col <= end.Column; col++)
                {
                    var cell = WorkSheet.Cells[row, col];
                    if (!String.IsNullOrEmpty(cell.Text) &&
                        cell.Text.StartsWith("[["))
                    {
                        TemplateFieldList.Add(cell.FullAddress, cell.Text);

                    }
                }
            }
        }
    }

}