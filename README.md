# ExcelToHtml
Excel To HTML Library and Console Application

# Getting Started

## Library 

~~~
FileInfo newFile = new FileInfo(fullPath);
var WorksheetHtml = new ExcelToHtml.ConvertToHtml(newFile);
string html = WorksheetHtml.ToHtml();
~~~

## Console App

~~~
ExcelToHtml.exe [Path]

Sample
ExcelToHtml.exe c:\excel.xlsx

Output
ExcelToHtml.exe c:\excel.xlsx.html
~~~

# Supported
-  Properties: Border, Text-align, background-color, font-weight, font-size, width, white-space
-  Horizontally Merged Cells
-  XLSX Files 

# Not Supported
- Horizontally merged cells
- Themes 
