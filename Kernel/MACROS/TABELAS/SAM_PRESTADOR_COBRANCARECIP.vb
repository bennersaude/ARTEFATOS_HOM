'HASH: F010E2C11702BB82ADE6E1B00EB8594F
'Macro SAM_PRESTADOR_COBRANCARECIP
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT CONVENIORECIPROCIDADE ")
	SQL.Add("  FROM SAM_PRESTADOR ")
	SQL.Add(" WHERE HANDLE = :HANDLE")

	SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	SQL.Active = True

	If (SQL.FieldByName("CONVENIORECIPROCIDADE").AsString <> "S") Then
		bsShowMessage("Prestador não está cadastrado como Convênio de Reciprocidade. Não é permitido incluir registros.", "E")
		Set SQL = Nothing
		CurrentQuery.Cancel
		RefreshNodesWithTable("SAM_PRESTADOR_COBRANCARECIP")
		Exit Sub
	End If

	Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim INTERFACE As Object
	Dim LINHA As String
	Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")

	LINHA = INTERFACE.Vigencia(CurrentSystem, "SAM_PRESTADOR_COBRANCARECIP", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", "")

	If (LINHA <> "") Then
		CanContinue = False
		bsShowMessage(LINHA, "E")
	End If

	Set INTERFACE = Nothing
End Sub
