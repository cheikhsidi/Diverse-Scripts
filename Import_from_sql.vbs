Sub ImportTable(sheet as String, query )
    On Error GoTo ErrorHandler
    Dim Cn As ADODB.Connection
    Dim Server_Name As Variant
    Dim Database_Name As Variant
    Dim User_ID As Variant
    Dim Password As Variant
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    'Server_Name = InputBox("Please Enter Server Name")
    'Database_Name = InputBox("Please Enter the Database Name")
    User_ID = InputBox("Please Enter your User Name")
    Password = InputBox("Please Enter your Password")
    Server_Name = "FRPBI.DATABASE.WINDOWS.NET" ' Enter your server name here
    Database_Name = "FRP_EDW" ' Enter your  database name here
    'SQLStr = "SELECT top 10 * FROM frps.Company" ' Enter your SQL here

    Set Cn = New ADODB.Connection
    Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"
    
    rs.Open query, Cn, adOpenStatic

    Dim Fileds As String
    Dim iCols As Integer

    With Worksheets(sheet)
    .Range("A2:Z10000").ClearContents
    For iCols = 0 To rs.Fields.Count - 1
    .Cells(10, iCols + 1).Value = rs.Fields(iCols).Name
    Next
    .Range("A11").CopyFromRecordset rs
    End With

    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    
    Exit Sub
ErrorHandler:
    MsgBox "Connection failed, please Re-Try"
    Exit Sub
End Sub



