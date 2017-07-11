using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtml
{
    public static class Theme
    {
        public static string[] Default = new string[] {
"FFFFFFFF",// REAL COLORS START FROM 1, Extra FF to match format
"FF000000",
"FFFFFFFF",
"FFFF0000",
"FF00FF00",
"FF0000FF",
"FFFFFF00",
"FFFF00FF",
"FF00FFFF",
"FF800000",
"FF008000",
"FF000080",
"FF808000",
"FF800080",
"FF008080",
"FFC0C0C0",
"FF808080",
"FF9999FF",
"FF993366",
"FFFFFFCC",
"FFCCFFFF",
"FF660066",
"FFFF8080",
"FF0066CC",
"FFCCCCFF",
"FF000080",
"FFFF00FF",
"FFFFFF00",
"FF00FFFF",
"FF800080",
"FF800000",
"FF008080",
"FF0000FF",
"FF00CCFF",
"FFCCFFFF",
"FFCCFFCC",
"FFFFFF99",
"FF99CCFF",
"FFFF99CC",
"FFCC99FF",
"FFFFCC99",
"FF3366FF",
"FF33CCCC",
"FF99CC00",
"FFFFCC00",
"FFFF9900",
"FFFF6600",
"FF666699",
"FF969696",
"FF003366",
"FF339966",
"FF003300",
"FF333300",
"FF993300",
"FF993366",
"FF333399",
"FF333333",
};
    }
}
//Macro To get all theme colors
//Sub colors56()
//'57 colors, 0 to 56
//  Application.ScreenUpdating = False
//  Application.Calculation = xlCalculationManual   'pre XL97 xlManual
//Dim i As Long
//Dim str0 As String, str As String
//For i = 0 To 56
//  Cells(i + 1, 1).Interior.ColorIndex = i
//  Cells(i + 1, 1).Value = "[Color " & i & "]"
//  Cells(i + 1, 2).Font.ColorIndex = i
//  Cells(i + 1, 2).Value = "[Color " & i & "]"
//  str0 = Right("000000" & Hex(Cells(i + 1, 1).Interior.Color), 6)
//  'Excel shows nibbles in reverse order so make it as RGB
//  str = Right(str0, 2) & Mid(str0, 3, 2) & Left(str0, 2)
//  'generating 2 columns in the HTML table
//  Cells(i + 1, 3) = """FF" & str & ""","
//Next i
//done:
//  Application.Calculation = xlCalculationAutomatic  'pre XL97 xlAutomatic
//  Application.ScreenUpdating = True
//End Sub
