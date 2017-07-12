using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtml
{
    public static  class Theme
    {


        public static Dictionary<string, string> DefaultTheme = new Dictionary<string, string>();

        public static void Init() {
            DefaultTheme.Add("BLACK", "#000000");
            DefaultTheme.Add("NAVY", "#000080");
            DefaultTheme.Add("BLUE", "#0000ff");
            DefaultTheme.Add("SLATE BLUE", "#007fff");
            DefaultTheme.Add("GREEN", "#008000");
            DefaultTheme.Add("TEAL", "#008080");
            DefaultTheme.Add("LIME", "#00ff00");
            DefaultTheme.Add("AQUA or CYAN", "#00ffff");
            DefaultTheme.Add("STEEL BLUE", "#236b8e");
            DefaultTheme.Add("SEA GREEN", "#238e6b");
            DefaultTheme.Add("MIDNIGHT BLUE", "#2f2f4f");
            DefaultTheme.Add("DARK GREEN", "#2f4f2f");
            DefaultTheme.Add("DARK SLATE GREY", "#2f4f4f");
            DefaultTheme.Add("MEDIUM BLUE", "#3232cc");
            DefaultTheme.Add("SKY BLUE", "#3299cc");
            DefaultTheme.Add("LIME GREEN", "#32cc32");
            DefaultTheme.Add("MEDIUM AQUAMARINE", "#32cc99");
            DefaultTheme.Add("CORNFLOWER BLUE", "#42426f");
            DefaultTheme.Add("INDIAN RED", "#4f2f2f");
            DefaultTheme.Add("VIOLET", "#4f2f4f");
            DefaultTheme.Add("DARK OLIVE GREEN", "#4f4f2f");
            DefaultTheme.Add("CADET BLUE", "#5f9f9f");
            DefaultTheme.Add("DARK SLATE BLUE", "#6b238e");
            DefaultTheme.Add("SALMON", "#6f4242");
            DefaultTheme.Add("DARK TURQUOISE", "#7093db");
            DefaultTheme.Add("AQUAMARINE", "#70db93");
            DefaultTheme.Add("MEDIUM TURQUOISE", "#70dbdb");
            DefaultTheme.Add("MEDIUM SLATE BLUE", "#7f00ff");
            DefaultTheme.Add("MEDIUM SPRING GREEN", "#7fff00");
            DefaultTheme.Add("MAROON", "#800000");
            DefaultTheme.Add("PURPLE", "#800080");
            DefaultTheme.Add("OLIVE", "#808000");
            DefaultTheme.Add("GREY", "#808080");
            DefaultTheme.Add("FIREBRICK", "#8e2323");
            DefaultTheme.Add("SIENNA", "#8e6b23");
            DefaultTheme.Add("LIGHT STEEL BLUE", "#8f8fbc");
            DefaultTheme.Add("PALE GREEN", "#8fbc8f");
            DefaultTheme.Add("MEDIUM ORCHID", "#9370db");
            DefaultTheme.Add("GREEN YELLOW", "#93db70");
            DefaultTheme.Add("DARK ORCHID", "#9932cc");
            DefaultTheme.Add("YELLOW GREEN", "#99cc32");
            DefaultTheme.Add("BLUE VIOLET", "#9f5f9f");
            DefaultTheme.Add("KHAKI", "#9f9f5f");
            DefaultTheme.Add("BROWN", "#A52A2A");
            DefaultTheme.Add("LIGHT GREY", "#a8a8a8");
            DefaultTheme.Add("TURQUOISE", "#adeaea");
            DefaultTheme.Add("PINK", "#bc8f8f");
            DefaultTheme.Add("LIGHT BLUE", "#bfd8d8");
            DefaultTheme.Add("SILVER", "#c0c0c0");
            DefaultTheme.Add("ORANGE RED", "#cc3232");
            DefaultTheme.Add("VIOLET RED", "#cc3299");
            DefaultTheme.Add("GOLD", "#cc7f32");
            DefaultTheme.Add("THISTLE", "#d8bfd8");
            DefaultTheme.Add("WHEAT", "#d8d8bf");
            DefaultTheme.Add("MEDIUM VIOLET RED", "#db7093");
            DefaultTheme.Add("ORCHID", "#db70db");
            DefaultTheme.Add("GOLDENROD", "#dbdb70");
            DefaultTheme.Add("PLUM", "#eaadea");
            DefaultTheme.Add("MEDIUM GOLDENROD", "#eaeaad");
            DefaultTheme.Add("RED", "#ff0000");
            DefaultTheme.Add("FUCHSIA", "#ff00ff");
            DefaultTheme.Add("ORANGE", "#ff7f00");
            DefaultTheme.Add("YELLOW", "#ffff00");
            DefaultTheme.Add("WHITE", "#ffffff");

        }

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
