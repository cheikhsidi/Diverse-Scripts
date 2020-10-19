Sub Reorganize_columns()
‘ Reorganize Columns Macro
‘ Description: Reorganize columns in Excel based on column headerDim v As Variant, x As Variant, findfield As Variant
Dim oCell As Range
Dim iNum As Long
v = Array(“First Name”, “Middle Name”, “Last Name”, “Date of Birth”, “Phone Number”, “Address”, “City”, “State”, “Postal (ZIP) Code”, “Country”)
For x = LBound(v) To UBound(v)
findfield = v(x)
iNum = iNum + 1
Set oCell = ActiveSheet.Rows(1).Find(What:=findfield, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)If Not oCell.Column = iNum Then
Columns(oCell.Column).Cut
Columns(iNum).Insert Shift:=xlToRight
End If
Next x
End Sub
