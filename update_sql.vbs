  
Sub UpdateTable(ByVal s As String, query As String)
    On Error GoTo ErrorHandler
    Dim Cn As ADODB.Connection
    Dim Server_Name As Variant
    Dim Database_Name As Variant
    Dim User_ID As Variant
    Dim Password As Variant
    'Dim SQLStr As String
   
    Dim iRowNo As Integer
    Dim Column1, Column2, Column3 As String
    Dim NRows As Integer

   
    User_ID = Trim(ThisWorkbook.Sheets("sheet2").Range("A1")) ' InputBox("Please Enter your User Name")
    Password = Trim(ThisWorkbook.Sheets("sheet2").Range("A2")) ' InputBox("Please Enter your Password")
  
    Server_Name = "FRPBI.DATABASE.WINDOWS.NET" ' Enter your server name here
    Database_Name = "FRP_EDW" ' Enter your  database name here
 

    'Set rngName = ActiveCell
    Set Cn = New ADODB.Connection
    
  
    With Sheets(s)
            
        'Open a connection to SQL Server
        Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
        ";Uid=" & User_ID & ";Pwd=" & Password & ";"
         
        'Skip the header row
        iRowNo = 4
            
        'Loop until empty cell in CustomerId
        Do Until .Cells(iRowNo, 1) = ""
            DataSource = Trim(.Cells(iRowNo, 1).value)
            MatchType = Trim(.Cells(iRowNo, 2).value)
            OldCode = Trim(.Cells(iRowNo, 3).value)
            NewCode = Trim(.Cells(iRowNo, 4).value)
            Name = Trim(.Cells(iRowNo, 5).value)
            NAICCode = Trim(.Cells(iRowNo, 6).value)
            Address1 = Trim(.Cells(iRowNo, 7).value)
            'MatchType = Trim(.Cells(iRowNo, 8))
            
         
            'Generate and execute sql statement to import the excel rows to SQL Server table
            Cn.Execute " Insert into frps.Company_map (DataSourceSK, MatchType, OldCode, NewCodeName, Code, NAICCode, Address1) values ('" & DataSourceSK & "', '" & MatchType & "', '" & OldCode & "','" & NewCode & "', '" & Name & "', '" & NAICCode & "','" & Address1 & "'); "

            Cn.Execute " IF EXISTS (SELECT * FROM frps.Company_map WHERE DataSourceSK ='" & DataSourceSK & "' and OldCode = '" & OldCode & "' and NewSCode = '" & NewCode & "') UPDATE frps.Company_map SET (Name = '" & Name & "', Code = '" & OldCode & "', FRPSCode = '" & Mapping & "', FRPSName = '" & FRPSName & "')" & _
            "WHERE DataSourceSK ='" & DataSource & "' and LegacyCode = '" & LegacyCode & "' and FRPSCode = '" & FRPSCode & "'" & _
            "Else insert into frps.Company_map (Name, Code, FRPSCode, FRPSName, NAICCode, Address1) values ('" & LegacyName & "','" & LegacyCode & "','" & FRPSCode & "', '" & FRPSName & "', '" & LegacyNAIC & "','" & Address1 & "'); "
 


            iRowNo = iRowNo + 1
        Loop
        'NRow = Cn.ExecuteNonQuery()
        'MsgBox NRow
        MsgBox "Updated succussefuly."
            
        Cn.Close
        Set Cn = Nothing
             
    End With

    Exit Sub
ErrorHandler:
    MsgBox "Update failed, please Re-Try"
    Exit Sub
 
End Sub

----------------------------------------------------------------------------------------------
Sub UpdateData()
Const adopenForward As Long = 0
Const adLockReadOnly As Long = 1
Const adCmdText As Long = 1
Dim oConn As Object
Dim oRS As Object
Dim sConnect As String
Dim sSQL As String
Dim ary

sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "c:\bob.mdb"

sSQL = "SELECT * From Contacts"
Set oRS = CreateObject("ADODB.Recordset")
oRS.Open sSQL, sConnect, adOpenForwardOnly, _
adLockReadOnly, adCmdText

' Check to make sure we received data.
If oRS.EOF Then
MsgBox "No records returned.", vbCritical
Else
sSQL = "UPDATE Contacts " & _
" SET Phone = 'None' " & _
"WHERE FirstName = 'Bob' AND LastNAme = 'Phillips'"
oRS.ActiveConnection.Execute sSQL

sSQL = "SELECT * From Contacts"
oRS.ActiveConnection.Execute sSQL
ary = oRS.getrows
MsgBox ary(0, 0) & " " & ary(1, 0) & ", " & ary(2, 0)
End If


oRS.Close
Set oRS = Nothing
End Sub