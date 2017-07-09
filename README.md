# ExcelToHtml
Excel To HTML Library and Console Application

# Getting Started

## Library  - Nuget Package 

Available Nuget Package https://www.nuget.org/packages/ExcelToHtml

~~~
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml = new ExcelToHtml.ConvertToHtml(newFile);
string html = WorksheetHtml.ToHtml();
~~~

## Console Application

~~~
ExcelToHtml.exe [Path]

Sample
ExcelToHtml.exe c:\excel.xlsx

Output
ExcelToHtml.exe c:\excel.xlsx.html
~~~

# List of Features
- Support for .xlsx format (Excel) 
- Excel Properties: Border, Text-align, background-color, font-weight, font-size, width, white-space
- Horizontal Merged Cells
- Injection safe


# List of Limitations
- Vertical merged cells
- Themes 
