Sub DropButtonClick()
    'Populate control.
    Me.cboClass.AddItem "Amphibian"
    Me.cboClass.AddItem "Bird"
    Me.cboClass.AddItem "Fish"
    Me.cboClass.AddItem "Mammal"
    Me.cboClass.AddItem "Reptile"
End Sub

Private Sub cboConservationStatus_DropButtonClick()
    'Populate control.
    Me.cboConservationStatus.AddItem "Endangered"
    Me.cboConservationStatus.AddItem "Extirpated"
    Me.cboConservationStatus.AddItem "Historic"
    Me.cboConservationStatus.AddItem "Special concern"
    Me.cboConservationStatus.AddItem "Stable"
    Me.cboConservationStatus.AddItem "Threatened"
    Me.cboConservationStatus.AddItem "WAP"
End Sub
Private Sub cboSex_DropButtonClick()
    'Populate control.
    Me.cboSex.AddItem "Female"
    Me.cboSex.AddItem "Male"
End Sub
Private Sub cmdAdd_Click()
    'Copy input values to sheet.
    Dim lRow As Long
    Dim ws As Worksheet
    Set ws = Worksheets("Animals")
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    With ws
        .Cells(lRow, 1).Value = Me.cboClass.Value
        .Cells(lRow, 2).Value = Me.txtGivenName.Value
        .Cells(lRow, 3).Value = Me.txtTagNumber.Value
        .Cells(lRow, 4).Value = Me.txtSpecies.Value
        .Cells(lRow, 5).Value = Me.cboSex.Value
        .Cells(lRow, 6).Value = Me.cboConservationStatus.Value
        .Cells(lRow, 7).Value = Me.txtComment.Value
    End With
    'Clear input controls.
    Me.cboClass.Value = ""
    Me.txtGivenName.Value = ""
    Me.txtTagNumber.Value = ""
    Me.txtSpecies.Value = ""
    Me.cboSex.Value = ""
    Me.cboConservationStatus.Value = ""
    Me.txtComment.Value = ""
End Sub
Private Sub cmdClose_Click()
    'Close UserForm.
    Unload Me
End Sub
