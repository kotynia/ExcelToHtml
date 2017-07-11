# ExcelToHtml
Excel To HTML Library and Console Application

# List of Features
- Convert Excel to HTML
	- Support for .xlsx format (Microsoft Office 2007+) 
	- Excel Properties: Border, Text-align, background-color(*), font-weight, font-size, width, white-space
	- Horizontal Merged Cells
	- Injection safe
- Optional INPUT/OUTPUT dataset (see Yaml File Format)
- Support of Excel functions in Formulas  ( https://epplus.codeplex.com/wikipage?title=Supported%20Functions&referringTitle=Documentation )

# Getting Started

## ExcelToHtml as a Library 

Download Nuget Package https://www.nuget.org/packages/ExcelToHtml

```c#
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml =  new ExcelToHtml.ToHtml(ExcelFile);


//Optional Get Set Cells
Dictionary<string, string> InputOutput = new Dictionary<string, string>();
InputOutput.Add("A1", "Hello World");  					//set hello world
InputOutput.Add("A2", "=2+1");  						//set formula
InputOutput.Add("[[TemplateField]]", "HelloTemplate");  //FillTempalte Field
InputOutput.Add(".A2", null);  							//Output value form A2
var output = WorksheetHtml.GetSetCells(InputOutput);	//Output


string html = WorksheetHtml.Convert();
```

## ExcelToHtml as a Console Application 

Download Latest Release https://github.com/marcinKotynia/ExcelToHtml/releases

```bat
ExcelToHtml.console.exe [Path]

Sample
ExcelToHtml.console.exe c:\myExcelFile.xlsx

Output
ExcelToHtml.console.exe c:\myExcelFile.xlsx.html

Optional
ExcelToHtml.exe c:\myExcelFile.xlsx.yaml
```

## Yaml File Format

Optional you can put file with data for example myExcelFile.xlsx.yaml
This file will provide values for Converter

Samples:

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

# List of Unsupported Features
- Vertical merged cells
- System Background Colors and Themes
- Charts

## *System Background Colors and Themes
Getting background color for a cell i really challenging.
There are 4 different scenarios which require different aproach

1. Themes
2. System Colors 
3. Manual selection from color picker (should be also system color)
4. Selecting other color than system manually from colorpicker.

Unfortunettly at the moment ExcelToHtml works with scenario 4.
workaround for other three is  manually switch color from color palette to manual colors.
Option could be to create macro for that.
Hope in near future to implement native parser and extend library.

Next step will be implement
https://stackoverflow.com/questions/10756206/getting-cell-backgroundcolor-in-excel-with-open-xml-2-0
https://simpleooxml.codeplex.com/ (as a extension class SpreadsheetReader does not exists in OpenXML 2.5 )


## Some technology remarks that could help you do even more :)
- xlsx file format is zip file with embeded xml files (https://en.wikipedia.org/wiki/Office_Open_XML )
- Libraries
 - EPPlus http://epplus.codeplex.com/ used currently 
 - ClosedXml https://github.com/ClosedXML/ClosedXML very active 
 - https://www.nuget.org/packages/DocumentFormat.OpenXml 
 - https://simpleooxml.codeplex.com/ 