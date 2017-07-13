# ExcelToHtml
Excel To HTML Library and Console Application

# List of Features
- Convert Excel to HTML
	- Support for .xlsx format (Microsoft Office 2007+) 
	- Excel Properties: Border, Text-align, background-color,color,font-weight, font-size, width, white-space
	- Horizontal Merged Cells
	- Hidden Rows and columns
	- Comments
	- Injection safe
- Optional INPUT/OUTPUT dataset (see Yaml File Format)
- Support for Functions  ( https://epplus.codeplex.com/wikipage?title=Supported%20Functions&referringTitle=Documentation )

## Road Map
- List<> to Excel Range
- Better handling Border Colors


# Getting Started

## ExcelToHtml as a Library, Nuget Package https://www.nuget.org/packages/ExcelToHtml

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);

//Optional set custom style for table
//WorksheetHtml.TableStyle =" " ; default "border-collapse: collapse;font-family: helvetica, arial, sans-serif;";

//Optional Get Set Cells
//Dictionary<string, string> InputOutput = new Dictionary<string, string>();
//InputOutput.Add("A1", "Hello World");  			//set hello world
//InputOutput.Add("A2", "=2+1");  			//set formula
//InputOutput.Add("[[TemplateField]]", "HelloTemplate");  //FillTempalte Field
//InputOutput.Add(".A2", null);  				//Output value form A2
//var output = WorksheetHtml.GetSetCells(InputOutput);	//Output


string html = WorksheetHtml.Convert();
```

## ExcelToHtml as a Console Application, Download https://github.com/marcinKotynia/ExcelToHtml/releases

How to use:

```bat
ExcelToHtml.console.exe [Path]

Sample
ExcelToHtml.console.exe c:\myExcelFile.xlsx

Output
ExcelToHtml.console.exe c:\myExcelFile.xlsx.html

Optional
ExcelToHtml.exe c:\myExcelFile.xlsx.yaml
```

# Getting Pro

## Parameters from Text file (Yaml)

Optional you can put file with data for example myExcelFile.xlsx.yaml

```yaml
# Set cell to 8
A3: 8
# Set cell to "Sample Text Value"
A4: Sample Text Value
# Set formula  , Formula must start with  = (equal)
A5: =A2+A3
# Instead of cell address you can set cell value in template for [[templatefield]]  and use from code
'[[templatefield]]': Hello Template field
# Output value , Value in yaml file (or dictionary) will be updated to calculated value at the end
.A5: 15
```



# Technical

## List of Unsupported Features
- Vertical merged cells
- Charts
- Images

## Colors and Themes
Getting color for a font, background is really challenging.
There are 3 different scenarios 

1. Themes (Supported only default theme)
2. System Colors with Index (supported)
3. RGB colors (supported)


This script will convert background color to rgb colors if you use custom theme.

```vb
Sub SheetBackgroundColorsToRgb()

Application.ScreenUpdating = False

'iterate
    For Each Cell In ActiveSheet.UsedRange.Cells
        'If Cell.Interior.Color > 0 Then
        'RGB
        Dim colorVal As Variant
        colorVal = Cell.Interior.Color
        Cell.Interior.Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        'End If
    Next
    
Application.ScreenUpdating = True

End Sub
```

## Some technology remarks that could help you do even more :)
- xlsx file format is zip file with embeded xml files (https://en.wikipedia.org/wiki/Office_Open_XML )
- Libraries
	- EPPlus http://epplus.codeplex.com/ used currently 
	- ClosedXml https://github.com/ClosedXML/ClosedXML very active 
	- https://www.nuget.org/packages/DocumentFormat.OpenXml 
	- https://simpleooxml.codeplex.com/ 
	
