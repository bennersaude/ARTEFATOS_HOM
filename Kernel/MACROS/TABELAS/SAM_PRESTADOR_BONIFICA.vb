'HASH: F9CB8B1B182B003D29F168543E898192
'Macro: SAM_PRESTADOR_BONIFICA
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull And CurrentQuery.FieldByName("DATAFINAL").AsDateTime < ServerDate Then
    DATAFINAL.ReadOnly = True
  Else
    DATAFINAL.ReadOnly = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String

	If CurrentQuery.FieldByName("BONUSBASEPF").AsInteger <= 0 And _
	   CurrentQuery.FieldByName("BONUSBASEBONIFICACAO").AsInteger <= 0 Then
		CanContinue = False
		bsShowMessage("Um dos percentuais deve ser maior que zero", "E")
		Exit Sub
	End If

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_BONIFICA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", "")

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	Dim vQuery As BPesquisa

	Set vQuery = NewQuery
	vQuery.Clear
	vQuery.Add("SELECT *")
	vQuery.Add("       FROM SAM_PRESTADOR_BONIFICA_EVENTO")
	vQuery.Add("WHERE PRESTADORBONIFICA = :PRESTADORBONIFICA ")
	vQuery.ParamByName("PRESTADORBONIFICA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	vQuery.Active = True

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	ElseIf Not vQuery.EOF Then
		bsShowMessage("Existem Eventos cadastrados nesta Bonificação. Exclusão não permitida", "E")
		CanContinue = False
	End If

	vQuery.Active = False
	Set vQuery = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMEssage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
