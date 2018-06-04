'HASH: CC64371768AC32929A1B2EFA36CA6025
'Macro: SAM_PRECOGERAL_ACOMODACAO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Set interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO|SAM_TIPOGRAU.DESCRICAO"
	vCriterio = ""
	vCampos = "Código do Grau|Descrição|Tipo do Grau"
	vHandle = interface.Exec(CurrentSystem, "SAM_GRAU|SAM_TIPOGRAU[SAM_TIPOGRAU.HANDLE = SAM_GRAU.TIPOGRAU]", vColunas, 2, vCampos, vCriterio, "Graus de Atuação", True, "")

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
	End If

	Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T4377" Then
			CONVENIO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim interface As Object
	Dim DataI As String
	Dim DataF As String
	Dim Linha As String
	Dim Condicao As String

	Condicao = " AND ACOMODACAO = " + CurrentQuery.FieldByName("ACOMODACAO").AsString

	If CurrentQuery.FieldByName("GRAU").IsNull Then
		Condicao = Condicao + " AND GRAU IS NULL"
	Else
		Condicao = Condicao + " AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString
	End If

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		Condicao = Condicao + " AND CONVENIO IS NULL"
	Else
		Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Set interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = interface.Vigencia(CurrentSystem, "SAM_PRECOGERAL_ACOMODACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ACOMODACAO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set interface = Nothing
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
End Sub
