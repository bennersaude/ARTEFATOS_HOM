'HASH: 472DDA628FF395DD0887117CFA19539E
'Macro: SAM_PRESTADOR_HORARIO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("HORAINICIAL").Value > CurrentQuery.FieldByName("HORAFINAL").Value Then
		CanContinue = False
		bsShowMessage("Hora final nao pode ser menor que a hora inicial!", "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("DOMINGO").Value = "N" Then
		If CurrentQuery.FieldByName("SEGUNDA").Value = "N" Then
			If CurrentQuery.FieldByName("TERCA").Value = "N" Then
				If CurrentQuery.FieldByName("QUARTA").Value = "N" Then
					If CurrentQuery.FieldByName("QUINTA").Value = "N" Then
						If CurrentQuery.FieldByName("SEXTA").Value = "N" Then
							If CurrentQuery.FieldByName("SABADO").Value = "N" Then
								CanContinue = False
								bsShowMessage("Horário inválido - deve-se marcar o(s) dia(s) da semana!", "E")
								Exit Sub
							End If
						End If
					End If
				End If
			End If
		End If
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
