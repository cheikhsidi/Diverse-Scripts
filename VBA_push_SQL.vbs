  
Sub UpdateTable()
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

    
    'Server_Name = InputBox("Please Enter Server Name")
    'Database_Name = InputBox("Please Enter the Database Name")
    'User_ID = InputBox("Please Enter your User Name")
    'Password = InputBox("Please Enter your Password")
    User_ID = Trim(ThisWorkbook.Sheets("sheet2").Range("A1")) ' InputBox("Please Enter your User Name")
    Password = Trim(ThisWorkbook.Sheets("sheet2").Range("A2")) ' InputBox("Please Enter your Password")
  
    Server_Name = "FRPBI.DATABASE.WINDOWS.NET" ' Enter your server name here
    Database_Name = "FRP_EDW" ' Enter your  database name here
 

    'Set rngName = ActiveCell
    Set Cn = New ADODB.Connection
    
  
    With Sheets("Sheet1")
            
        'Open a connection to SQL Server
        Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
        ";Uid=" & User_ID & ";Pwd=" & Password & ";"
         
        'Skip the header row
        iRowNo = 11
            
        'Loop until empty cell in CustomerId
        Do Until .Cells(iRowNo, 1) = ""
            DataSource = CInt(Trim(.Cells(iRowNo, 1)))
            LegacyName = Trim(.Cells(iRowNo, 2))
            LegacyCode = Trim(.Cells(iRowNo, 3))
            FRPSCode = Trim(.Cells(iRowNo, 4))
            FRPSName = Trim(.Cells(iRowNo, 5))
            LegacyNAIC = Trim(.Cells(iRowNo, 6))
            Address1 = Trim(.Cells(iRowNo, 7))
            MatchType = Trim(.Cells(iRowNo, 8))
            
         
            'Generate and execute sql statement to import the excel rows to SQL Server table
            Cn.Execute " Insert into frps.Company_map (Name, Code, FRPSCode, FRPSName, NAICCode, Address1) values ('" & LegacyName & "','" & LegacyCode & "','" & FRPSCode & "', '" & FRPSName & "', '" & LegacyNAIC & "','" & Address1 & "'); "
 
            iRowNo = iRowNo + 1
        Loop
        'NRow = Cn.ExecuteNonQuery()
        'MsgBox NRow
        MsgBox "FRPS CompanyMap imported."
            
        Cn.Close
        Set conn = Nothing
             
    End With

    Exit Sub
ErrorHandler:
    MsgBox "Update failed, please Re-Try"
    Exit Sub
 
End Sub


------------------------------------------------------
Private Sub CommandButton1_Click()


    'On Error GoTo ErrorHandler
    Dim Cn As ADODB.Connection
    Dim Server_Name As Variant
    Dim Database_Name As Variant
    Dim User_ID As Variant
    Dim Password As Variant
    
    Dim DataSource As Variant
    Dim MatchType As String
    Dim OldCode As String
    Dim NewCode As String
    Dim Name As String
    Dim NAICCode As Variant
    Dim Address1 As String
   
    Dim iRowNo As Integer
   
    Dim sheets As Variant
    Dim sheet As Variant
    Dim ws As Worksheet
    Dim n As String
   
   
    User_ID = Trim(ThisWorkbook.sheets("sheet2").Range("A1")) ' InputBox("Please Enter your User Name")
    Password = Trim(ThisWorkbook.sheets("sheet2").Range("A2")) ' InputBox("Please Enter your Password")
  
    Server_Name = "FRPBI.DATABASE.WINDOWS.NET" ' Enter your server name here
    Database_Name = "FRP_EDW" ' Enter your  database name here
 

    'Set rngName = ActiveCell
    Set Cn = New ADODB.Connection
    
              
        'Open a connection to SQL Server
        Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
        ";Uid=" & User_ID & ";Pwd=" & Password & ";"
         
        'truncate all tables befor the insert to avaoid duplication
       
        Cn.Execute "truncate table etl.Company_mapStg"
        Cn.Execute "truncate table etl.PolicyLineType_mapStg"
        Cn.Execute "truncate table etl.PolicyLineStatus_mapStg"
        Cn.Execute "truncate table etl.Department_mapStg"
        Cn.Execute "truncate table etl.Branch_mapStg"
        Cn.Execute "truncate table etl.Vendor_mapStg"
        
        
    'shs = Array("Company", "PolicyLineType", "PolicyLineStatus", "Dertment", "Branch", "Vendor", "Employee")
    'tables = Array("etl.Company_mapStg", "etl.PolicyLineType_mapStg", "etl.PolicyLineStatus_mapStg", "etl.Department_mapStg", "etl.Branch_mapStg", "etl.Vendor_mapStg", "Employee")
  
    For Each ws In ThisWorkbook.Worksheets
        n = ws.Name
        If n <> "Sheet2" And n <> "Queries" And n <> "Administration" And n <> "Employee" Then
     
        'T = tables(i)
        iRowNo = 4
        
        'Loop until empty cell in CustomerId
        Do Until ws.Cells(iRowNo, 1).Value = ""
            DataSource = ws.Cells(iRowNo, 1).Value
            MatchType = ws.Cells(iRowNo, 2).Value
            OldCode = ws.Cells(iRowNo, 3).Value
            NewCode = ws.Cells(iRowNo, 4).Value
            Name = CStr(ws.Cells(iRowNo, 5).Value)
            NAICCode = ws.Cells(iRowNo, 6).Value
            Address1 = ws.Cells(iRowNo, 7).Value
    
         
            'Generate and execute sql statement to import the excel rows to SQL Server table
                       
            If n = "Company" Then
                Cn.Execute " Insert into etl.Company_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "','" & NAICCode & "','" & Address1 & "')"

            ElseIf n = "Vendor" Then
               Cn.Execute " Insert into etl.Vendor_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "','" & Address1 & "')"
     
            ElseIf n = "Department" Then
                Cn.Execute " Insert into etl.Department_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')"
                
            ElseIf n = "Branch" Then
                Cn.Execute " Insert into etl.Branch_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')"
            ElseIf n = "PolicyLineType" Then
                Cn.Execute " Insert into etl.PolicyLineType_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')"
                
            ElseIf n = "PolicyLineStatus" Then
                Cn.Execute " Insert into etl.PolicyLineStatus_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')"
                
            End If
            iRowNo = iRowNo + 1
           
        Loop
        End If
        i = i + 1
        
        Next ws
        
        
        MsgBox "Insert Succuss."
            
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
----------------------------------------------------------------
Sub replace(ByVal s As String)
'replace single quote by adding another quote to escape
 shWorksheets(s).Columns("E").replace _
 What:="'", Replacement:="''", _
 SearchOrder:=xlByColumns, MatchCase:=True
End Sub


