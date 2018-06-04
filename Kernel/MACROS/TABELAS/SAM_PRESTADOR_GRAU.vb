'HASH: D747BE50FF3CBAA9E36380D320DC1BCC
'Macro: SAM_PRESTADOR_GRAU
'02/01/2001 -Alterado por Paulo Garcia Junior -liberacao para edição do registro atraves dos parametros gerais de prestador
'#Uses "*liberaRegraExcecao"
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If liberaRegraExcecao <>"" Then
		DATAINICIAL.ReadOnly = True
		DATAFINAL.ReadOnly = True
		PRESTADOR.ReadOnly = True
		REGRAEXCECAO.ReadOnly = True
		GRAU.ReadOnly = True
	Else
		DATAINICIAL.ReadOnly = False
		DATAFINAL.ReadOnly = False
		PRESTADOR.ReadOnly = False
		REGRAEXCECAO.ReadOnly = False
		GRAU.ReadOnly = False
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = liberaRegraExcecao

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = liberaRegraExcecao

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = liberaRegraExcecao

	If Msg <>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = "AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_GRAU", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) T")
	SQL.Add("  FROM SAM_PRESTADOR_GRAU X")
	SQL.Add(" WHERE X.PRESTADOR = :P")
	SQL.Add("   AND X.GRAU = :G")
	SQL.Add("   AND X.REGRAEXCECAO IN ('R','E')")

	SQL.ParamByName("P").Value = RecordHandleOfTable("SAM_PRESTADOR")
	SQL.ParamByName("G").Value = CurrentQuery.FieldByName("GRAU").AsInteger
	SQL.Active = True

	If SQL.FieldByName("T").AsInteger >1 Then
		CanContinue = False
		bsShowMessage("O Grau não pode ser registrado como Regra e exceção ao mesmo tempo", "E")
	End If
End Sub
