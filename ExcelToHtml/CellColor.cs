using ClosedXML.Excel;
using System;

namespace ExcelToHtml
{
    public static class colors
    {

        //closedxml experiment
        public static string getCellBackgroundColor(IXLWorksheet workSheet, string cellAddress)
        {

            IXLCell cell = workSheet.Cell(cellAddress);
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
                if (Theme.DefaultTheme.TryGetValue(cell.Style.Fill.BackgroundColor.ThemeColor.ToString(), out value))
                    return value;
                else
                    Console.WriteLine("Theme not supported {2} cell col:{0} row:{1}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.ThemeColor);
               
            }
            return string.Empty;
        }

        public static string getCellTextColor(IXLWorksheet workSheet, string cellAddress)
        {

            IXLCell cell = workSheet.Cell(cellAddress);
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
                if (Theme.DefaultTheme.TryGetValue(cell.Style.Font.FontColor.ThemeColor.ToString(), out value))
                    return value;
                else
                    Console.WriteLine("Theme not supported {2} cell col:{0} row:{1}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Font.FontColor.ThemeColor);

            }
            return string.Empty;
        }
    }
}