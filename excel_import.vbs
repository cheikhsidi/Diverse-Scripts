Function RangeToInsert(rRng As Range, table as string) As String

    Dim vaData As Variant
    Dim i As Long, j As Long
    Dim aReturn() As String
    Dim aCols() As String
    Dim aVals() As Variant
    'Const table as String
    Const sINSERT As String = "INSERT INTO" & table 
    Const sVAL As String = " VALUES "

    'Read in data
    vaData = rRng.Value

    'Create arrays
    ReDim aReturn(1 To UBound(vaData))
    ReDim aCols(1 To UBound(vaData, 2))
    ReDim aVals(1 To UBound(vaData, 2))

    'Fill column name array from first row
    For j = LBound(vaData, 2) To UBound(vaData, 2)
        aCols(j) = vaData(1, j)
    Next j

    'Go through the rest of the rows
    For i = LBound(vaData, 1) + 1 To UBound(vaData, 1)

        'Fill a temporary array
        For j = LBound(vaData, 2) To UBound(vaData, 2)
            aVals(j) = vaData(i, j)
        Next j

        'Build the string into the main array
        aReturn(i) = sINSERT & "(" & Join(aCols, ",") & ")" & sVAL & "(" & Join(aVals, ",") & ");"
    Next i

    RangeToInsert = Join(aReturn, vbNewLine)

End Function