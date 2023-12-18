Private Sub CommandButton3_Click()
Dim LastRow As Long
LastRow = WorksheetFunction.CountA(Sheets("source1").Range("A:A"))

If UsrFrm.Cmbbox1.Value = "" Or UsrFrm.Txtbox2.Value = "" Or UsrFrm.Cmbox3.Value = "" Or UsrFrm.Txtbox4.Value = "" Or UsrFrm.TextBox4.Value = "" Or UsrFrm.TextBox3.Value = "" Or UsrFrm.TextBox2.Value = "" Then
MsgBox "أدخل البيانات كاملة"

Else

Sheets("source1").Cells(LastRow + 1, 1).Value = UsrFrm.Cmbbox1.Value
Sheets("source1").Cells(LastRow + 1, 2).Value = UsrFrm.Txtbox2.Value
Sheets("source1").Cells(LastRow + 1, 3).Value = UsrFrm.Cmbox3.Value
Sheets("source1").Cells(LastRow + 1, 8).Value = UsrFrm.Txtbox4.Value
Sheets("source1").Cells(LastRow + 1, 5).Value = UsrFrm.TextBox1.Value
Sheets("source1").Cells(LastRow + 1, 6).Value = UsrFrm.TextBox2.Value
Sheets("source1").Cells(LastRow + 1, 7).Value = UsrFrm.TextBox3.Value
Sheets("source1").Cells(LastRow + 1, 4).Value = UsrFrm.TextBox4.Value


UsrFrm.Cmbbox1.Value = ""
UsrFrm.Txtbox2.Value = ""
UsrFrm.Cmbox3.Value = ""
UsrFrm.Txtbox4.Value = ""
UsrFrm.TextBox4.Value = ""
UsrFrm.TextBox3.Value = ""
UsrFrm.TextBox1.Value = ""
UsrFrm.TextBox2.Value = ""
End If

End Sub

Private Sub btnshowdata_Click()
psw = InputBox("Please Enter The Password")
If psw = 1234 Then
Application.Visible = True
Sheets("source1").Visible = True
Sheets("source1").Activate
UsrFrm.Hide

Else
MsgBox "WRONG PASSWORD ! ASK DEVELOPER HALA"
End If
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Terminate()
ActiveWorkbook.Save
Application.Quit
End Sub