'HASH: 96CCE4744DAA523CF63A89E234B7308B
 
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	CanContinue = False

	If CurrentQuery.FieldByName("IDADEFINAL").AsInteger < CurrentQuery.FieldByName("IDADEINICIAL").AsInteger Then
		MsgBox("A idade final não pode ser menor que a idade inicial!")
		IDADEFINAL.SetFocus
		Exit Sub
	End If

	CanContinue = True
End Sub
