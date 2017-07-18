# ExcelToHtml.dll , ExcelToHtml.console.exe 
Excel To HTML Library and Console Application

# List of Features (1.3)

- Convert Excel to HTML
	- Support for .xlsx format (Microsoft Office 2007+) 
	- Excel Properties: Border,border collor, Text-align, background-color, color, font-weight, font-size, width, white-space
	- Horizontal Merged Cells
	- Hidden Rows and columns
	- Comments
	- Injection safe
- Support for Functions  ( https://epplus.codeplex.com/wikipage?title=Supported%20Functions&referringTitle=Documentation )
- Calculation Engine
- Merge object, Json, REST API and excel template, convert to html

# Getting Started

## ExcelToHtml.dll, Nuget Package https://www.nuget.org/packages/ExcelToHtml

Basic Convert excel to HTML

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml = new ExcelToHtml.ToHtml(ExcelFile);
string html = WorksheetHtml.Convert();
```

ExcelToHtml as calculation engine InputOutput

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);

//Optional Get Set Cells
Dictionary<string, string> InputOutput = new Dictionary<string, string>();
InputOutput.Add("A1", "Hello World");  			//set hello world
InputOutput.Add("A2", "=2+1");  			//set formula
InputOutput.Add("[[TemplateField]]", "HelloTemplate");  //FillTempalte Field
InputOutput.Add(".A2", null);  				//Output value form A2
var output = WorksheetHtml.GetSetCells(InputOutput);	//Output

string html = WorksheetHtml.Convert();
```

Merge REST API data and excel template, get html

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DebugMode = true;
WorksheetHtml.DataFromUrl("http://nflarrest.com/api/v1/crime");
string html = WorksheetHtml.GetHtml();
```

Merge object and excel template, get html

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DataFromObject(object); 
string html = WorksheetHtml.GetHtml();
```

Merge json and excel template, get html 

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DataFromJson(string); 
string html = WorksheetHtml.GetHtml();
```

Merge json and excel template, get Excel

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);
WorksheetHtml.DataFromJson(string); 
string html = WorksheetHtml.GetBytes();
```

## ExcelToHtml.console.exe, Download https://github.com/marcinKotynia/ExcelToHtml/releases

How to use:

```bat
@Echo Get HTML
exceltohtml.console.exe Test1.xlsx
@Echo Output Test1.xlsx.html

@Echo Get XLSX
exceltohtml.console.exe -t=Test1.xlsx -output=xlsx
@Echo Output Test1.xlsx.xlsx

@Echo Get html PAGE with https://www.w3schools.com/w3css/default.asp stylesheet
exceltohtml.console.exe -t=Test1.xlsx -output=htmlw3css
@Echo Output Test1.xlsx.html

@Echo Get PDF !! Warning require https://wkhtmltopdf.org/downloads.html  
exceltohtml.console.exe -t=Test1.xlsx -output=pdf
@Echo  Output Test1.xlsx.pdf


@Echo Convert and Merge data from url (REST API) 
exceltohtml.console.exe -t=Test1.xlsx  -data=http://nflarrest.com/api/v1/crime
@Echo Output Test1.xlsx.html

```

# Getting Pro

## Parameters from text file (Yaml)

Optional parameters for file with data Test1.xlsx.yaml

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

## Merging Object or JSON 

ExcelToHtml can use any object or JSON.
Example will merge data from public REST API into Excel

```bat
exceltohtml.console.exe -t=Test1.xlsx  -data=http://nflarrest.com/api/v1/crime
```

```JSON
[
{"Category":"DUI","arrest_count":"187"},
{"Category":"Drugs","arrest_count":"93"},
{"Category":"Domestic violence","arrest_count":"84"},
{"Category":"Assault","arrest_count":"65"},
{"Category":"Disorderly conduct","arrest_count":"38"},
{"Category":"Gun","arrest_count":"26"}
]
```

To use this data you can put following template fields

- *[[d[0][!].Category]]* will render to list of elements from Category
- *[[d[0][!].arrest_count]]* will render to list of elements from arrest_count
- *[[d[0][0].arrest_count]]* will render to element value 187

# Evolution 

Excel ScreenShot
![Image](Test/Test1%20ExcelToHtml.png?raw=true)

Version 1.0.0 result
![1.0.0](Test/Test1%20ExcelToHtml1.0.0.png?raw=true)

Version 1.1.0 result
![1.1.0](/Test/Test1%20ExcelToHtml1.1.0.png?raw=true)

Version 1.2.0 result
![1.2.0](/Test/Test1%20ExcelToHtml1.2.0.png?raw=true)

Version 1.3.0 result
![1.3.0](/Test/Test1%20ExcelToHtml1.3.0.png?raw=true)

# Technical Appendix

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


This script will convert background color and font color to rgb colors if you use custom theme
and colours. To use 

1. open file in Excel 
2. Alt+F11 
3. Paste and Run code using F5

Result: Colors (background,borders,font) will be converted to RGB colors

```vb
Sub SheetBackgroundColorsToRgb()

Application.ScreenUpdating = False

    For Each Cell In ActiveSheet.UsedRange.Cells
    
		'Background
        Dim colorVal As Variant
        colorVal = Cell.Interior.Color
        Cell.Interior.Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        
        'Font color
        colorVal = Cell.Font.Color
        If (Not colorVal) Then
        Cell.Font.Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        'Borders     
        colorVal = Cell.Borders(xlEdgeBottom).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeBottom).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        colorVal = Cell.Borders(xlEdgeRight).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeRight).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        colorVal = Cell.Borders(xlEdgeTop).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeTop).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
        colorVal = Cell.Borders(xlEdgeLeft).Color
        If (Not colorVal) Then
        Cell.Borders(xlEdgeLeft).Color = RGB((colorVal Mod 256), ((colorVal \ 256) Mod 256), (colorVal \ 65536))
        End If
        
    Next
    
Application.ScreenUpdating = True

End Sub
```

## Some technology remarks that could help you do even more :)
- xlsx file format is zip file with embeded xml files (https://en.wikipedia.org/wiki/Office_Open_XML )
- Libraries thet will help you
	- EPPlus http://epplus.codeplex.com/ 
	- ClosedXml https://github.com/ClosedXML/ClosedXML 
	- Microsoft wrapper for handling openxml https://www.nuget.org/packages/DocumentFormat.OpenXml 
	
