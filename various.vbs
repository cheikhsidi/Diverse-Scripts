'Create new sheets
Sub newSh()
Dim w As Worksheet

 For Each ws In ThisWorkbook.Worksheets
  n = ws.Name
  If n <> "Sheet2" And n <> "Queries" And n <> "Administration" And n <> "Employee" And Not n Like "*_*" Then
  'ws.Copy After:=ThisWorkbook.sheets(n & "_L")
    'Set w = ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count))
    'w.Name = n & "_L"
    End If
    Next ws
End Sub
--------------------------------------------------------------------------------------------------------------------
'User Authentication
Private Sub CommandButton1_Click()
  Dim objTargetWorksheet As Worksheet

  If (TextBox1.Value = "John" And TextBox2.Value = "234") _
    Or (TextBox1.Value = "Amy" And TextBox2.Value = "345") _
    Or (TextBox1.Value = "Paul" And TextBox2.Value = "456") Then
    Me.Hide: Application.Visible = True

    For Each objTargetWorksheet In ActiveWorkbook.Worksheets
      If objTargetWorksheet.Name = TextBox1.Value Then
        objTargetWorksheet.Unprotect Password:=12345
      Else
        objTargetWorksheet.Protect Password:=12345, DrawingObjects:=True, Contents:=True, Scenarios:=True
      End If
    Next
  Else
    MsgBox "Please input the right user name and the right password"
  End If
End Sub

Private Sub CommandButton2_Click()
  ThisWorkbook.Application.Quit
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  ThisWorkbook.Application.Quit
End Sub
------------------------------------------------------------------------------------------
'to prevent a macro from beun run with password
Dim password As Variant
password = Application.InputBox("Enter Password", "Password Protected")

Select Case password
    Case Is = False
        'do nothing
    Case Is = "easy"
        Range("A1").Value = "This is secret code"
    Case Else
        MsgBox "Incorrect Password"
End Select 