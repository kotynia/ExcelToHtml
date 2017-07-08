using OfficeOpenXml;
using System;
using System.IO;
using System.Text;

namespace ExcelToHtml
{
    public class ConvertToHtml
    {
        //if not specified first used
        private string WorksheetName = String.Empty;
        ExcelPackage Excel;
        ExcelWorksheet WorkSheet;

        public string tableAtributeStyle = "border-collapse: collapse;font-family: helvetica, arial, sans-serif;";

        public ConvertToHtml(FileInfo excelFile)
        {
            if (!excelFile.Exists)
                throw new Exception(String.Format("File {0} Not Found", excelFile.FullName));

            Excel = new ExcelPackage(excelFile);
        }

        public string ToHtml(string WorkSheetName=null) {
            if (!string.IsNullOrEmpty(WorkSheetName))
                WorkSheet = Excel.Workbook.Worksheets[WorksheetName];
            else
                WorkSheet = Excel.Workbook.Worksheets[1];

            //GET DIMENSIONS
            var start = WorkSheet.Dimension.Start;
            var end = WorkSheet.Dimension.End;
            StringBuilder sb = new StringBuilder();

            sb.AppendLine(String.Format("<table style=\"{0}>\"",tableAtributeStyle));

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

            return sb.ToString();
        }
    }
}
