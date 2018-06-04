'HASH: 6B9FE918C0D0EEF5605A4E2A4DC47340
'#USES "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (DATAINICIAL.EditDate > DATAFINAL.EditDate) Then
		bsShowMessage("Data inicial não pode ser maior que a data final", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	Select Case CommandID
		Case ""
	End Select

End Sub
