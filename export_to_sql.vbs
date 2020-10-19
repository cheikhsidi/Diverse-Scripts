Sub UploadTable()
If MsgBox("You are about to upload/commit to the Database, do you want to continue?", vbYesNo) = vbNo Then Exit Sub
    'On Error GoTo ErrorHandler
    Dim Cn As ADODB.Connection
    
    Dim DataSource As String
    Dim MatchType As String
    Dim OldCode As String
    Dim NewCode As String
    Dim Name As String
    Dim NAICCode As String
    Dim Address1 As String
   
    Dim i, iRowNo As Integer
   
    Dim ws As Worksheet
    Dim n As String
    Dim recordsAffected As Long
    Dim StrSproc As String
    
    'Setting new Connection
    Set Cn = New ADODB.Connection
              
        'Open a connection to SQL Server
        Cn.Open connStr
         
        'Truncate all tables befor the insert to avaoid duplication
        Cn.Execute "truncate table etl.Company_mapStg"
        Cn.Execute "truncate table etl.mappingErrors_mapStg"
        Cn.Execute "truncate table etl.PolicyLineType_mapStg"
        Cn.Execute "truncate table etl.PolicyLineStatus_mapStg"
        Cn.Execute "truncate table etl.Department_mapStg"
        Cn.Execute "truncate table etl.Branch_mapStg"
        Cn.Execute "truncate table etl.Vendor_mapStg"
        Cn.Execute "truncate table etl.Broker_mapStg"
                
    'tables = Array("etl.Company_mapStg", "etl.PolicyLineType_mapStg", "etl.PolicyLineStatus_mapStg", "etl.Department_mapStg", "etl.Branch_mapStg", "etl.Vendor_mapStg", "Employee")
    'Looping through worsheets and retreiving data from sql
    For Each ws In ThisWorkbook.Worksheets
        n = ws.Name
        If n <> "Sheet2" And n <> "Queries" And n <> "Administration" And n <> "Employee" Then
        
        i = 0
        iRowNo = 4
        
        'Loop until empty cell in DataSource
        Do Until ws.Cells(iRowNo, 1).Value = ""
        If ws.Cells(iRowNo, 8).Value = 1 Or (ws.Cells(iRowNo, 2).Value <> "NoMatch" And ws.Cells(iRowNo, 2).Value <> "preMap") Then
            DataSource = ws.Cells(iRowNo, 1).Value
            MatchType = ws.Cells(iRowNo, 2).Value
            OldCode = ws.Cells(iRowNo, 3).Value
            NewCode = ws.Cells(iRowNo, 4).Value
            Name = CStr(ws.Cells(iRowNo, 5).Value)
            NAICCode = ws.Cells(iRowNo, 6).Value
            Address1 = ws.Cells(iRowNo, 7).Value
        
    
            'Generate and execute sql statement to import the excel rows to the target SQL Server table
                       
            If n = "Company" Then
                Cn.Execute " Insert into etl.Company_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "','" & NAICCode & "','" & Address1 & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("C13").Value = i
            
            ElseIf n = "Vendor" Then
               Cn.Execute " Insert into etl.Vendor_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "','" & Address1 & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("K24").Value = i
            
            ElseIf n = "Department" Then
                Cn.Execute " Insert into etl.Department_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("C24").Value = i
            
            ElseIf n = "Branch" Then
                Cn.Execute " Insert into etl.Branch_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("G24").Value = i
            
            ElseIf n = "PolicyLineType" Then
                Cn.Execute " Insert into etl.PolicyLineType_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("G13").Value = i
            
            ElseIf n = "PolicyLineStatus" Then
                Cn.Execute " Insert into etl.PolicyLineStatus_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("K13").Value = i
            
            ElseIf n = "Broker" Then
                Cn.Execute " Insert into etl.Broker_mapStg values ('" & DataSource & "','" & MatchType & "','" & OldCode & "', '" & NewCode & "', '" & Name & "')", recordsAffected
            i = i + recordsAffected
            Worksheets("Administration").Range("C36").Value = i
            End If
            
        End If
            iRowNo = iRowNo + 1
           
        Loop
        End If
        i = i + 1
       
        Next ws
        
        'Invoking the Stored Procedures query to reflect the Chnages into the excel
        'StrSproc = "set nocount on; EXEC FRPS.DataIntegrationUpdateFromExcel"
        StrSproc = "set nocount on; EXEC frps.DataIntegrationUpdateFromExcel_v2 0"
        Cn.Execute StrSproc, recordsAffected
        MsgBox "Insert Succuss."
            
        Cn.Close
        Set Cn = Nothing

' Error Handleing
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
