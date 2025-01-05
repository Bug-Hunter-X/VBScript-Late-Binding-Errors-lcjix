Late Binding: VBScript's late binding can cause runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where version mismatches or typos in object names can lead to unexpected failures.

Example:
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")

' Assume a typo in the method name
objectExcel.WorkBooks.Open "myFile.xls"
```
This might fail silently or throw a generic error message, making debugging difficult.