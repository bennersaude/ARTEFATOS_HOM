'HASH: 64CDB8C82968B073F3C84A922C841CE7
'Macro: SAM_PRESTADOR_CURRICULO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If Not CurrentQuery.FieldByName("DATACONCLUSAO").IsNull Then
		If CurrentQuery.FieldByName("DATACONCLUSAO").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
			bsShowMessage("Data da conclusão não pode ser menor que a data Inicial!", "E")
			CanContinue = False
		End If
	End If

	If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > ServerDate Then
		bsShowMessage("Data inicial não pode ser maior que a data atual!", "E")
		CanContinue = False
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
	Dim SQL As Object
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQL = NewQuery

	SQL.Add("SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = :P")

	SQL.ParamByName("P").Value = RecordHandleOfTable("SAM_PRESTADOR")
	SQL.Active = True

	If SQL.FieldByName("FISICAJURIDICA").AsInteger <> 1 Then
		CanContinue = False
		bsShowMessage("O registro do currículo destina-se somenente a prestadores pessoa física", "E")
	End If

	Set SQL = Nothing
End Sub
