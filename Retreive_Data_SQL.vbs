' Retreiving Data from SQL Server passing the the destination sheet name and the query as arguments

Sub ImportTable(ByVal s As String, query As String)
    'On Error GoTo ErrorHandler
    Dim Cn As ADODB.Connection
   
  
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
 
    Set Cn = New ADODB.Connection
    
    Cn.Open connStr
    rs.Open query, Cn, adOpenStatic

    'Dim Fileds As String
    'Dim iCols As Integer
    
    With Worksheets(s)
    '.ShowAllData
    .Range("A4:Z100000").ClearContents
    'For iCols = 0 To rs.Fields.Count - 1
    '.Cells(3, iCols + 1).Value = rs.Fields(iCols).name
    'Next
    .Range("A4").CopyFromRecordset rs
    .Range("H4:H100000").ClearContents
    End With
    

    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    
    
    Exit Sub
ErrorHandler:

    MsgBox "Error : " & Err & " " & Error(Err)
     If Not Cn Is Nothing Then
      Set Cn = Nothing
     End If
     If Not rs Is Nothing Then
      Set rs = Nothing
     End If
     
    Exit Sub
    
    
End Sub

Sub Clear_All_Filters_Range(ByVal ws As Worksheet)

  'To Clear All Fitlers use the ShowAllData method for
  'for the sheet.  Add error handling to bypass error if
  'no filters are applied.  Does not work for Tables.
  On Error Resume Next
    ws.ShowAllData
  On Error GoTo 0
  
End Sub

'Genrate the connection String
Function connStr() As Variant
     Dim Server_Name As String
     Dim Database_Name As String
     Dim User_ID As String
     Dim Password As String
     
     User_ID = Trim(ThisWorkbook.sheets("sheet2").Range("A1")) ' Please Enter your User Name
 
     Password = Trim(ThisWorkbook.sheets("sheet2").Range("A2")) ' Please Enter your Password
    
     Server_Name = "FRPBI.DATABASE.WINDOWS.NET" ' Enter your server name here
     Database_Name = "FRP_EDW" ' Enter your  database name here
     
     connStr = "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"
     
End Function



