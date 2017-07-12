using ClosedXML.Excel;
using System;

namespace ExcelToHtml
{
    public static class colors
    {

        //closedxml experiment
        public static string getColor(IXLWorksheet workSheet, string cellAddress)
        {

            IXLCell cell = workSheet.Cell(cellAddress);
            if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
            {
               // Console.WriteLine("color {0} {1} {2}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.Color.Name);
                return cell.Style.Fill.BackgroundColor.Color.ToHex();
            }
            else if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
            {
                if (cell.Style.Fill.BackgroundColor.Color.Name != "Transparent")
                {
                  
                      return cell.Style.Fill.BackgroundColor.Color.ToHex();
                }
            }
            else  //(cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)
            {
                
                string value = "";
                if (Theme.DefaultTheme.TryGetValue(cell.Style.Fill.BackgroundColor.ThemeColor.ToString(), out value))
                {
                    return value;
                }
                else
                {
                    Console.WriteLine("color not found {0}{1} {2}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.ThemeColor);
                }
                
            }

            return string.Empty;

        }





        //test iteration
        public static void GetColor111()
        {


            using (XLWorkbook workBook = new XLWorkbook(@"c:\git\ExcelToHtml\Test\test1.xlsx"))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                foreach (IXLRow row in workSheet.Rows())
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)

                            Console.WriteLine("color {0} {1} {2}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.Color.Name);
                        else if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Indexed)
                        {
                            if (cell.Style.Fill.BackgroundColor.Color.Name != "Transparent")
                                Console.WriteLine("index {0} {1} {2}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.Color.Name);
                        }
                        else  //(cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Theme)

                            Console.WriteLine("theme {0} {1} {2}", cell.Address.ColumnLetter, cell.Address.RowNumber, cell.Style.Fill.BackgroundColor.ThemeColor);
                    }
                }
            }


        }
    }
}