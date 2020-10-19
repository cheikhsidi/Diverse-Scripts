' Validation macro
Private Sub Worksheet_Change(ByVal Target As Range)

Dim Cn As ADODB.Connection
    Dim Server_Name As Variant
    Dim Database_Name As Variant
    Dim User_ID As Variant
    Dim Password As Variant
    Dim SQLStr As String
    Dim sheetrow As Integer
    Dim sheetcolumn As Integer
    Dim v As String
  
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    'Server_Name = InputBox("Please Enter Server Name")
    'Database_Name = InputBox("Please Enter the Database Name")
    'ThisWorkbook.Sheets("Sheet3").Activate
    User_ID = Trim(ThisWorkbook.Sheets("sheet2").Range("A1")) ' InputBox("Please Enter your User Name")
 
    Password = Trim(ThisWorkbook.Sheets("sheet2").Range("A2")) ' InputBox("Please Enter your Password")
    
    ThisWorkbook.Sheets("Sheet1").Activate
    Server_Name = "FRPBI.DATABASE.WINDOWS.NET" ' Enter your server name here
    Database_Name = "FRP_EDW" ' Enter your  database name here
    
    
 With Worksheets("Sheet1")
 
 'If Not Intersect(Target, .ComboBox10) Is Nothing Then
        v1 = .ComboBox10.Value
        
        Range("H10").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$10:$I$17934").AutoFilter Field:=8, Criteria1:="" & v1
        'End If
        
    If Not Application.Intersect(Range("I:I"), Range(Target.Address)) Is Nothing Then
 'If Not Intersect(Target, .Range("I11")) Is Nothing Then

  If Not Target.Value = "" Then 'Exit Sub
  
  'Else
    v = Target.Value
 
    SQLStr = "SELECT * FROM FRPS.CompanyMapping where FRPSCode = '" & v & "' " ' Enter your SQL here

    Set Cn = New ADODB.Connection
    Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"
    
    rs.Open SQLStr, Cn, adOpenStatic
    
    sheetrow = ActiveCell.Row
    sheetcolumn = ActiveCell.Column
    
    If rs.RecordCount = 0 Then
    
    Cells(sheetrow, sheetcolumn + 1).Value = 0
    
    Else
    Cells(sheetrow, sheetcolumn + 1).Value = 1
    
     End If
rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing


 End If


End If

 End With


End Sub

'Filter 
Sheet1.ComboBox10.Clear
    
    i = 0
    Do
    DoEvents
    i = i + 1
    item_in_review = TheSheet.Range("A" & i)
       
    If Len(item_in_review) > 0 Then Sheet1.ComboBox10.AddItem (item_in_review)
    
    Loop Until item_in_review = ""
    
    Sheet1.ComboBox10.ListIndex = 0
    

' Validation Handeler
    Private sub Worksheet_Change(ByVal Target As Range)

    Sheetrow = ActiveCell.Row
    sheetcolumn = ActiveCell.Column
    
 a = Application.WorksheetFunction.CountA(Sheet1.Range("A:A"))
    if Not Application.Intersect(Range("$A$10:$A$" & a), Range(Target.Address)) Is Nothing Then
    