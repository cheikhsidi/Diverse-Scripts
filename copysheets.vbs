'Copy only data from a sheet to another sheet
Sub cop()
Sheets("Companies").Range("A1:E212").Copy Destination:=Sheets("Sheet2").Range("A1:E212")
End Sub

'Copy a sheet into another workbook
Worksheets("Sheet1").Copy
With ActiveWorkbook 
     .SaveAs Filename:=Environ("TEMP") & "\New1.xlsx", FileFormat:=xlOpenXMLWorkbook
     .Close SaveChanges:=False
End With

'Copy Multiple sheets into new workbook
Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Copy
With ActiveWorkbook
     .SaveAs Filename:=Environ("TEMP") & "\New3.xlsx", FileFormat:=xlOpenXMLWorkbook 
     .Close SaveChanges:=False 
End With 