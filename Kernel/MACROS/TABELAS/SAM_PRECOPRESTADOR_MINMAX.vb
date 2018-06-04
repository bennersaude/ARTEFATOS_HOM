'HASH: 1C67E2CD82A0426DD2466103C76B4AD6
'Macro: SAM_PRECOPRESTADOR_MINMAX
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"

Public Function Verifica_ehAssociacao As Boolean
	Verifica_ehAssociacao = False

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT ASSOCIACAO FROM SAM_PRESTADOR WHERE HANDLE = :HPRESTADOR")

	SQL.ParamByName("HPRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
	SQL.Active = True

	If SQL.FieldByName("ASSOCIACAO").AsString = "N" Then
		CLASSEASSOCIADO.Visible = False
		Exit Function
	Else
		CLASSEASSOCIADO.Visible = True
	End If

	Set SQL = Nothing

	Verifica_ehAssociacao = True
End Function

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	If ShowPopup = False Then
		Exit Sub
	End If

	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraGrau(GRAU.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If Verifica_ehAssociacao Then
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("VALORMAXIMO").AsFloat <CurrentQuery.FieldByName("VALORMINIMO").AsFloat Then
		bsShowMessage("O valor máximo deve ser maior que o mínimo", "E")
		CanContinue = False
		Exit Sub
	End If

	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = Condicao + " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	Condicao = Condicao + " AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString

	If Not CurrentQuery.FieldByName("GRAU").IsNull Then
		Condicao = Condicao + " AND (GRAU IS NULL OR GRAU = " + CurrentQuery.FieldByName("GRAU").AsString + ")"
	End If

	If Verificaehassociacao Then
		Condicao = Condicao + " AND CLASSEASSOCIADO = '" + CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString + "'"
	End If

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOPRESTADOR_MINMAX", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
		Exit Sub
	End If

	Set Interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If Verifica_ehAssociacao Then
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If Verifica_ehAssociacao Then
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeScroll()
	If Verifica_ehAssociacao Then
		Exit Sub
	End If
End Sub

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

	SQL.Active = True

	If SQL.FieldByName("TOTAL").AsInteger = 1 Then
		SQL.Active = False

		SQL.Clear

		SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

		SQL.Active = True

		CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
	End If

	Set SQL = Nothing

	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
	End If
End Sub
