
Private Sub CommandButton2_Click()
    'Defining connection and workbooks and variables
    Dim wb As Workbook
    Dim Rng As Range
    Dim t As String
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    'Dim sArray() As Variant
    Dim rRng As Range
    
    'Insert statment variables
    Dim vaData As Variant
    Dim i As Long, j As Long
    Dim aReturn() As String
    Dim aCols() As Range
    Dim aVals() As Variant
    Dim DataSource As String
    Dim sINSERT As String
    Dim ws As Worksheet
    Dim strSearch As String

    'Open a connection to SQL Server
     Cn.Open connStr
    
    t = ThisWorkbook.Sheets("sheet1").Cells(6, 2).Value
    p = ThisWorkbook.Sheets("sheet1").Cells(3, 2).Value
    wn = ThisWorkbook.Sheets("sheet1").Cells(4, 2).Value
    sh = ThisWorkbook.Sheets("sheet1").Cells(5, 2).Value
    Set wb = Workbooks.Open(p & wn)
        
    'Set column_range = wb.Worksheets("ActivityCodes").Cells(i).EntireColumn
    With wb.Sheets("" & sh)
        LastRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
        lColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        MsgBox "Total Records :" & LastRow
        MsgBox "Total Columns :" & lColumn
        Set rRng = .Range(.Cells(1, 1), .Cells(LastRow, lColumn))
               
   'Creating an insert statement from the selected range
    sINSERT = "INSERT INTO " & t
    Const sVAL As String = " VALUES "

    'Read in data
    vaData = rRng.Value
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set w = wb.Sheets("" & sh)
        
    'Create arrays
    'Loop until empty cell in in target table
    lastR = ws.Range("C10").End(xlDown).Row
    MsgBox lastR
         
    ReDim aReturn(1 To UBound(vaData))
    ReDim aVals(10 To lastR)

    For i = LBound(vaData, 1) + 1 To UBound(vaData, 1)
        
        'Fill column name array from first row
        For j = 10 To lastR - 1
                
            strSearch = ws.Cells(j, 4).Value

            If Not strSearch = "" Then
            Set aCell = .Rows(1).Find(What:=strSearch, LookIn:=xlValues, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
            
             If InStr(LCase(strSearch), "date") Then
               If .Cells(i, aCell.Column).Text = "" Then
                    aVals(j) = ""

                Else: aVals(j) = Format(.Cells(i, aCell.Column).Text, "YYYY-MM-DD HH:MM")
                End If
             Else:
             'aVals(j) = .Cells(i, aCell.Column).Text
                aVals(j) = Replace(.Cells(i, aCell.Column).Text, "'", "''")
             End If
           
            'Number2Letter (aCell.Column)
            'aVals(j) = .Cells(i, aCell.Column).Text
            'aVals(j) = """& " & Number2Letter(aCell.Column) & i & " &"""
            ElseIf InStr(strSearch, """") Then
                aVals(j) = Format(strSearch, "YYYY-MM-DD HH:MM")
            Else:
            aVals(j) = ""
            End If
        Next j
      
        aReturn(i) = sINSERT & sVAL & "('" & Join(aVals, "','") & "');"
        'MsgBox aReturn(i)

        'Build the string into the main array insert statment
        'aReturn(i) = sINSERT & sVAL & "('" & Join(aVals, "','") & "');"
        
        
        
        Cn.Execute aReturn(i)
    Next i
        
    End With
   MsgBox "Inserted successfully", vbInformation

End Sub
