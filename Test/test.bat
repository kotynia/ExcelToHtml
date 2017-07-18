@Echo Convert to HTML
exceltohtml.console.exe Test1.xlsx

@Echo Convert and Merge data from url (REST API) 
exceltohtml.console.exe -t=Test1.xlsx  -data=http://nflarrest.com/api/v1/crime -outputpath=test11

@Echo Convert and Merge data from url (REST API) 
exceltohtml.console.exe -t=Test1.xlsx  -data=http://nflarrest.com/api/v1/crime -output=xlsx -outputpath=test12

@Echo Convert and Merge data from url (REST API) 
exceltohtml.console.exe -t=Test1.xlsx  -data=http://nflarrest.com/api/v1/crime -output=pdf -outputpath=test13

@Echo Convert and Merge data from url (REST API) 
exceltohtml.console.exe -t=Test1.xlsx  -data=http://nflarrest.com/api/v1/crime -output=htmlw3css -outputpath=test14

PAUSE
