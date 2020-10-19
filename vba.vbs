sub CreateSheets()

    Dim wb As Workbook: Set wb =
    Dim strName As String: strName =
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(Type:=xlWorksheet)
    With ws
        .Name = strName
    End With


End Sub