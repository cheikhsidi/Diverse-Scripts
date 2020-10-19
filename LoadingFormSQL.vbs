Sub main()
 'On Error GoTo ErrorHandler
 Dim Name As String
 Dim ws As Worksheet
 Dim n, q As String
  
 With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .CutCopyMode = False
End With
 
 'Validation, in case someone hit it accidentally
 If MsgBox("You are about to refresh/load which will overwrite any change you have made, do you want to continue?", vbYesNo) = vbNo Then Exit Sub

 'Looping through All the sheets to load data from sql into the target sheet
 For Each ws In ThisWorkbook.Worksheets
  Name = ws.Name
  If Name <> "Sheet2" And Name <> "Queries" And Name <> "Administration" And Name <> "Employee" And Not Name Like "*_*" Then
     ws.Unprotect Password:="12345"
     
     ' Spliting the DataSource, to retreive only the number to pass it to the where clause
     n = Trim(Split("" & Worksheets("Administration").selectDS.Value, "-")(0))
     'Reading the queries from the Queries worksheet
     q = Application.WorksheetFunction.VLookup(Name, Worksheets("Queries").Range("A:B"), 2, False)
     
     'Clear all the filters
     Call Clear_All_Filters_Range(ws)
     
     'Calling the Subroutine that import data from sql
     Call ImportTable(Name, q & n & "order by MatchType, Name")
     'escaping single quote in the name column by doubling it
    Call replace(ws, "E")
  'escaping single quote in the address column as well
    If Name = "Company" Then
        Call replace(ws, "G")
    End If
     
    'locking the worksheet except the column New code
     ws.Protect Password:="12345", contents:=True, userinterfaceonly:=True, AllowFiltering:=True
     ws.Range("D4:D100000").Locked = False
    
    End If
      
    Next ws
    'Activating the Admin sheet
    ThisWorkbook.sheets("Administration").Activate
    MsgBox "Succussfully Retreived"
    
    'Handling Errors
       Exit Sub
ErrorHandler:
    MsgBox "Connection failed, please Re-Try"
    Exit Sub
    
    
 
End Sub
 

Sub replace(ByVal ws As Worksheet, s As String)
'replace single quote by adding another quote
 ws.Columns(s).replace _
 what:="'", Replacement:="''", _
 SearchOrder:=xlByColumns, MatchCase:=True
End Sub
