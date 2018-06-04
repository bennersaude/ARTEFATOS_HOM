'HASH: A1986C7D69C2230099784875A509545B
'Macro: SAM_PRECOPRESTADOR_SL
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaUS(TABELAUS.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAUS").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterEdit()
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
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String
	'SMS 49152 - Anderson Lonardoni
	'Esta verificação foi tirada do BeforeInsert e colocada no
	'BeforePost para que, no caso de Inserção, já existam valores
	'no CurrentQuery e para funcionar com o Integrator
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
	'SMS 49152 - Fim

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Condicao = " AND PORTESALA = " + CurrentQuery.FieldByName("PORTESALA").AsString
	Condicao = Condicao + " AND CLASSEASSOCIADO = '" + CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString + "'"

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOPRESTADOR_SL", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
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
