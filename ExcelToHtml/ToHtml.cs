using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using OfficeOpenXml.Style;
using ClosedXML.Excel; //used from develop branch (fix with loading template )

namespace ExcelToHtml
{
    public class ToHtml
    {

        public string TableStyle = "border-collapse: collapse;font-family: helvetica, arial, sans-serif;";
        public static Dictionary<string, string> Theme = new Dictionary<string, string>();

        /// If not specified first used
        private string WorksheetName = String.Empty;
        ExcelPackage Excel;
        ExcelWorksheet WorkSheet;
        IXLWorksheet closedWorksheet; //closedxml  to get valid colors
        private Dictionary<string, string> TemplateFieldList;
        private Dictionary<string,string> cellStyles;
        

        public ToHtml(FileInfo excelFile, string WorkSheetName = null)
        {
            if (!excelFile.Exists)
                throw new Exception(String.Format("File {0} Not Found", excelFile.FullName));

            Excel = new ExcelPackage(excelFile);

            XLWorkbook workBook = new XLWorkbook(excelFile.FullName); //closedxml only temporary to get valid colors


            if (!string.IsNullOrEmpty(WorkSheetName))
            {
                WorkSheet = Excel.Workbook.Worksheets[WorksheetName];
                closedWorksheet = workBook.Worksheet(WorksheetName);//closedxml only temporary to get valid colors
            }
            else
            {
                WorkSheet = Excel.Workbook.Worksheets[1];
                closedWorksheet = workBook.Worksheet(1);//closedxml only temporary to get valid colors
            }
            Theme  = ExcelToHtml.Theme.Init();

        }

        /// <summary>
        /// Render HTML
        /// </summary>
        /// <returns>Html</returns>
        public string Execute()
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
                if (!WorkSheet.Row(row).Hidden)
                {
                    sb.AppendLine("<tr>");
                    for (int col = start.Column; col <= end.Column; col++)
                    {

                        if (!WorkSheet.Column(col).Hidden)
                        {
                            var d = WorkSheet.Cells[row, col];

                            int merged = 0;

                            if (d.Merge) //row is merged
                                merged = d.Worksheet.SelectedRange[WorkSheet.MergedCells[row, col]].Columns;

                            //11 default font size
                            var x = ProcessCellStyle(WorkSheet.Cells[row, col], WorkSheet.Column(col).Width, 11, merged);
                            sb.AppendLine(x);
                            if (d.Merge)
                                col += (merged - 1);
                        }
                    }
                    sb.AppendLine("</tr>");
                }
            }

            sb.AppendLine("</table>");
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            Console.WriteLine("Total time {0}ms", elapsedMs);

            return string.Format("<table  style=\"{0}>\" data-eth-ms=\"{1}\" data-eth-date=\"{2}\">{3}</table>",
                TableStyle, elapsedMs, DateTime.Now, sb.ToString());
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

                    TemplateMap(); //One time

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
        private void TemplateMap()
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

        private string ProcessCellStyle(ExcelRange input, double Width = -1, int FontSize = 11, int ColSpan = 0)
        {
            cellStyles = new Dictionary<string,string>();

            StringBuilder sb = new StringBuilder();

            //Border
            PropertyToStyle("border-top", input.Style.Border.Top);
            PropertyToStyle("border-right", input.Style.Border.Right);
            PropertyToStyle("border-bottom", input.Style.Border.Bottom);
            PropertyToStyle("border-left", input.Style.Border.Left);

            //Align
            PropertyToStyle("text-align", input.Style.HorizontalAlignment.ToString(), "General");

            //Colors
            //Not properly implemented in Epplus  using ClosedXML
            PropertyToStyle("background-color", getCellBackgroundColor( input.Address));        
            PropertyToStyle("color", getCellTextColor( input.Address));


            PropertyToStyle("font-weight", input.Style.Font.Bold == true ? "bold" : "");
            PropertyToStyle("font-size", input.Style.Font.Size.ToString(), "11");
            PropertyToStyle("width", Convert.ToInt16(Width * 10));

            PropertyToStyle("white-space", input.Style.WrapText == false ? "no-wrap" : "");


            string value = input.Text;
            if (string.IsNullOrEmpty(value))
                value = "&nbsp;";
            else
                value = System.Net.WebUtility.HtmlEncode(value);



        string comment   = (input.Comment!= null && input.Comment.Text !="" ) ? ("title=\"" + input.Comment.Text+"\"") : string.Empty ;


            if (ColSpan > 0)
                sb.AppendFormat("<td style=\"{0}\" eth-cell=\"{1}\" colspan=\"{2}\" {4} >{3}</td>",
                    string.Join(";", cellStyles.Select(x => x.Key + ":" + x.Value)), input.Address, ColSpan, value,comment);
            else
                sb.AppendFormat("<td style=\"{0}\" eth-cell=\"{1}\" {3} >{2}</td>", 
                    string.Join(";",cellStyles.Select(x => x.Key + ":" + x.Value)), input.Address, value, comment);

            return sb.ToString();

        }

        private void PropertyToStyle(string cssproperty, object value, string defaultommit = "",string cellAddress = "")
        {
            if (value == null)
                return;

            string cssItem = string.Empty;

            //borders
            if (value.GetType() == typeof(ExcelBorderItem))
            {
                var temp = (ExcelBorderItem)value;

                if (temp.Style == ExcelBorderStyle.None)
                    return;
                else if (temp.Style == ExcelBorderStyle.Thin)
                    cssItem = "solid 1px";
                else if (temp.Style == ExcelBorderStyle.Hair)
                    cssItem = "solid 1px";
                else if (temp.Style == ExcelBorderStyle.Thick)
                    cssItem = "solid 1px";
                else if (temp.Style == ExcelBorderStyle.Dashed)
                    cssItem = "dashed 1px";
                else if (temp.Style == ExcelBorderStyle.Dotted)
                    cssItem = "dotted 1px";
                else
                    cssItem = "solid 2px";

                //string color = getCellBorderColor(cellAddress,temp);

                if (!string.IsNullOrEmpty(temp.Color.Theme))  //no idea how to get proper theme color
                    cssItem += " #000";
                else if (!string.IsNullOrEmpty(temp.Color.Rgb)) //Remov First FF
                    cssItem += "#" + temp.Color.Rgb.Remove(0, 2);
                else
                    cssItem += " #000"; //default color if not defined

                cellStyles.Add(cssproperty, cssItem);
                return;
            }
            else
            {
                cssItem = value.ToString();
            }

            if (cssItem != defaultommit)
            {
                if (cssproperty.Contains("size") || cssproperty.Contains("width"))
                {
                    cellStyles.Add( cssproperty, cssItem.Replace(",",".") + "px");
                }
                else if (cssproperty.Contains("color")) //Remove First FF
                {
                    cellStyles.Add( cssproperty, "#" + cssItem.Remove(0, 2));
                }
                else
                    cellStyles.Add( cssproperty, cssItem);
            }
        }

        //private string getCellBorderColor(string cellAddress, ExcelBorderItem border)
        //{

        //    IXLCell cell = closedWorksheet.Cell(cellAddress);
        //    if (border.ColorType == XLColorType.Color)
        //    {
        //        return border.Color.ToHex();
        //    }
        //    else if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
        //    {
        //        if (cell.Style.Fill.BackgroundColor.Color.Name != "Transparent")
        //            return cell.Style.Fill.BackgroundColor.Color.ToHex();

        //    }
        //    else  //(cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
        //    {
        //        string value = "";
        //        if (Theme.TryGetValue(cell.Style.Fill.BackgroundColor.ThemeColor.ToString(), out value))
        //            return value;
        //        else
        //            Console.WriteLine("Theme not found {2} cell:{0}{1}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.ThemeColor);

        //    }
        //    return string.Empty;
        //}

        private  string getCellBackgroundColor( string cellAddress)
        {

            IXLCell cell = closedWorksheet.Cell(cellAddress);
            if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
            {
                return cell.Style.Fill.BackgroundColor.Color.ToHex();
            }
            else if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
            {
                if (cell.Style.Fill.BackgroundColor.Color.Name != "Transparent")
                    return cell.Style.Fill.BackgroundColor.Color.ToHex();

            }
            else  //(cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
            {
                string value = "";
                if (Theme.TryGetValue(cell.Style.Fill.BackgroundColor.ThemeColor.ToString(), out value))
                    return value;
                else
                    Console.WriteLine("Theme not found {2} cell:{0}{1}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.ThemeColor);

            }
            return string.Empty;
        }

        private  string getCellTextColor(string cellAddress)
        {

            IXLCell cell = closedWorksheet.Cell(cellAddress);
            if (cell.Style.Font.FontColor.ColorType == XLColorType.Color)
            {
                return cell.Style.Font.FontColor.Color.ToHex();
            }
            else if (cell.Style.Font.FontColor.ColorType == XLColorType.Indexed)
            {
                if (cell.Style.Font.FontColor.Color.Name != "Transparent")
                    return cell.Style.Font.FontColor.Color.ToHex();

            }
            else  //(cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
            {
                string value = "";
                if (Theme.TryGetValue(cell.Style.Font.FontColor.ThemeColor.ToString(), out value))
                    return value;
                else
                    Console.WriteLine("Theme not found {2} cell:{0}{1}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Font.FontColor.ThemeColor);

            }
            return string.Empty;
        }

    }
}