using System;

public class Class1
{

        public static PatternFill GetCellPatternFill(Cell theCell, SpreadsheetDocument document)
        {
            WorkbookStylesPart styles = SpreadsheetReader.GetWorkbookStyles(document);

            int cellStyleIndex;
            if (theCell.StyleIndex == null) // I think (from testing) if the StyleIndex is null
            {                               // then this means use cell style index 0.
                cellStyleIndex = 0;           // However I did not found it in the open xml 
            }                               // specification.
            else
            {
                cellStyleIndex = (int)theCell.StyleIndex.Value;
            }

            CellFormat cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[cellStyleIndex];

            Fill fill = (Fill)styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value];
            return fill.PatternFill;
        }

        private static void PrintColorType(SpreadsheetDocument sd, DocumentFormat.OpenXml.Spreadsheet.ColorType ct)
        {
            if (ct.Auto != null)
            {
                Console.Out.WriteLine("System auto color");
            }

            if (ct.Rgb != null)
            {
                Console.Out.WriteLine("RGB value -> {0}", ct.Rgb.Value);
            }

            if (ct.Indexed != null)
            {
                Console.Out.WriteLine("Indexed color -> {0}", ct.Indexed.Value);

                //IndexedColors ic = (IndexedColors)styles.Stylesheet.Colors.IndexedColors.ChildElements[(int)bgc.Indexed.Value];         
            }

            if (ct.Theme != null)
            {
                Console.Out.WriteLine("Theme -> {0}", ct.Theme.Value);

                Color2Type c2t = (Color2Type)sd.WorkbookPart.ThemePart.Theme.ThemeElements.ColorScheme.ChildElements[(int)ct.Theme.Value];

                Console.Out.WriteLine("RGB color model hex -> {0}", c2t.RgbColorModelHex.Val);
            }

            if (ct.Tint != null)
            {
                Console.Out.WriteLine("Tint value -> {0}", ct.Tint.Value);
            }
        }

        static void ReadAllBackgroundColors()
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("c:\git\ExcelToHtml\Test\test1.xlsx", false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                {
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    foreach (Row r in sheetData.Elements<Row>())
                    {
                        foreach (Cell c in r.Elements<Cell>())
                        {
                            Console.Out.WriteLine("----------------");
                            PatternFill pf = GetCellPatternFill(c, spreadsheetDocument);

                            Console.Out.WriteLine("Pattern fill type -> {0}", pf.PatternType.Value);

                            if (pf.PatternType == PatternValues.None)
                            {
                                Console.Out.WriteLine("No fill color specified");
                                continue;
                            }

                            Console.Out.WriteLine("Summary foreground color:");
                            PrintColorType(spreadsheetDocument, pf.ForegroundColor);
                            Console.Out.WriteLine("Summary background color:");
                            PrintColorType(spreadsheetDocument, pf.BackgroundColor);
                        }
                    }
                }
            }
        }
    
    
}
