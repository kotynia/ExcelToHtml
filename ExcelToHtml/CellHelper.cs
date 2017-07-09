using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text;


class CellHelper
{

    static List<string> items;

    
    public static string Process(ExcelRange input, double Width = -1, int FontSize = 11, int ColSpan = 0)
    {
        items = new List<string>();

        StringBuilder sb = new StringBuilder();




        //Border
        PropertyToCss("border-top", input.Style.Border.Top);
        PropertyToCss("border-right", input.Style.Border.Right);
        PropertyToCss("border-bottom", input.Style.Border.Bottom);
        PropertyToCss("border-left", input.Style.Border.Left);

        //Align
        PropertyToCss("text-align", input.Style.HorizontalAlignment.ToString(), "General");

        //Colors
        PropertyToCss("background-color", input.Style.Fill.BackgroundColor.Rgb);
        PropertyToCss("color", input.Style.Font.Color.Rgb);
        PropertyToCss("font-weight", input.Style.Font.Bold == true ? "bold" : "");
        PropertyToCss("font-size", input.Style.Font.Size.ToString(), "11");
        PropertyToCss("width", Convert.ToInt16(Width * 10));

        PropertyToCss("white-space", input.Style.WrapText == false ? "no-wrap" : "");

        // if (items.Count() > 0) //apply visual style

        string value = input.Text;
        if (string.IsNullOrEmpty(value))
            value = "&nbsp;";
        else
            value = System.Net.WebUtility.HtmlEncode(value);

        

        if (ColSpan > 0)
            sb.AppendFormat("<td style=\"{0}\" colspan=\"{1}\">{2}</td>", String.Join<string>(String.Empty, items), ColSpan, value);
        else
            sb.AppendFormat("<td style=\"{0}\">{1}</td>", String.Join<string>(String.Empty, items), value);

        return sb.ToString();

    }


    static void PropertyToCss(string cssproperty, object value, string defaultommit = "")
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
                cssItem = "dotted 2px";


            if (!string.IsNullOrEmpty(temp.Color.Theme))  //no idea how to get proper theme color
                cssItem += " #000";
            else if (!string.IsNullOrEmpty(temp.Color.Rgb)) //Remov First FF
                cssItem += "#" + temp.Color.Rgb.Remove(0, 2);
            else
                cssItem += " #000"; //default color if not defined

            items.Add(string.Format("{0}:{1};", cssproperty, cssItem));
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
                items.Add(string.Format("{0}:{1}px;", cssproperty, cssItem));
            }
            else if (cssproperty.Contains("color")) //Remove First FF
            {
                items.Add(string.Format("{0}:#{1};", cssproperty, cssItem.Remove(0, 2)));
            }
            else
                items.Add(string.Format("{0}:{1};", cssproperty, cssItem));
        }
    }

}