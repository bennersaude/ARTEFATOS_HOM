'HASH: CF3B19F96B1A23120269D782624E3488
'Macro: SAM_PRESTADOR_EQUIPAMENTO
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_SAM_PRESTADOR_EQUIPAMENTO" Then
			PRESTADOR.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("ANO").AsInteger > Year(ServerDate) Then
		CanContinue = False
		bsShowMessage("Ano de fabricacao nao pode ser maior que o ano atual!", "E")
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
