
Private Sub Refresh()
 Dim Name As String
 Dim ws As Worksheet
 Dim q As String
 
If MsgBox("You are about to upload/commit to the Database, do you want to continue?", vbYesNo) = vbNo Then Exit Sub
    On Error GoTo ErrorHandler

'Looping through All the sheets to load data from sql into the target sheet
 For Each ws In ThisWorkbook.Worksheets
  Name = ws.Name
  If Name <> "Sheet2" And Name <> "Queries" And Name <> "Administration" And Name <> "Employee" And Name Like "*_*" Then
     ws.Unprotect Password:="12345"
     
     'Reading the queries from the Queries worksheet
     q = Application.WorksheetFunction.VLookup(Name, Worksheets("Queries").Range("A:B"), 2, False)
     
     'Clear all the filters
     Call Clear_All_Filters_Range(ws)
     
     'Calling the Subroutine that import data from sql
     Call ImportTable(Name, q)
     
    'locking the worksheet except the column New code
     ws.Protect Password:="12345", contents:=True, userinterfaceonly:=True, AllowFiltering:=True
     ws.Visible = False
    End If
      
  Next ws


       Exit Sub
ErrorHandler:
    MsgBox "Connection failed, please Re-Try"
    Exit Sub
End Sub

'A Sreach Macro
Sub Search()
    Dim rng As Range
    Dim val As String
    Dim lastC As Integer
    Dim n As Integer
    Dim FoundCell As Range
    Dim FirstAddr As String
    
       
    n = 11
    val = Trim("" & ThisWorkbook.sheets("Company").ComboBox2.Value)
    If val = "" Then
    Exit Sub
    End If
    
    Set FoundCell = Worksheets("Company_L").Range("C:C").Find(what:=val, LookAt:=xlPart)
    If Not FoundCell Is Nothing Then
        ThisWorkbook.sheets("Company").Range("K10:K100").ClearContents
        ThisWorkbook.sheets("Company").Range("K10") = FoundCell.Offset(0, -1) & " - " & FoundCell
        FirstAddr = FoundCell.Address
    Else
        MsgBox val & " not found"
        ThisWorkbook.sheets("Company").Range("K10:K100").ClearContents
        ThisWorkbook.sheets("Company").Range("K10") = "NA"
        Exit Sub
    End If
        Do Until FoundCell Is Nothing
        Debug.Print FoundCell.Address
    Set FoundCell = Worksheets("Company_L").Range("C:C").FindNext(after:=FoundCell)
    ThisWorkbook.sheets("Company").Range("K" & n) = FoundCell.Offset(0, -1) & " - " & FoundCell
    If FoundCell.Address = FirstAddr Then
        Exit Do
    End If
    
    n = n + 1
Loop
   
End Sub