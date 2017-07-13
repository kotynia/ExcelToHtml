using System.Collections.Generic;

namespace ExcelToHtml
{
    public static  class Theme
    {
       

        public static Dictionary<string, string> Init() {

            //Default Excel color Theme
            Dictionary<string, string> DefaultTheme = new Dictionary<string, string>();
            DefaultTheme.Add("Background1", "FFFFFFFF");
            DefaultTheme.Add("Background2", "FFE7E6E6");
            DefaultTheme.Add("Text1", "FF000000");
            DefaultTheme.Add("Text2", "FF44546A");
            DefaultTheme.Add("Accent1", "FF4472C4");
            DefaultTheme.Add("Accent2", "FFED7D31"); 
            DefaultTheme.Add("Accent3", "FFA5A5A5");
            DefaultTheme.Add("Accent4", "FFFFC000"); 
            DefaultTheme.Add("Accent5", "FF5B9BD5"); 
            DefaultTheme.Add("Accent6", "FF70AD47"); 
            return DefaultTheme;
        }

    }
}