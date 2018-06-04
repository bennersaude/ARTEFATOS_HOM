'HASH: BA83F3AD5291841C03AE7FE7EA3CD0C1
'#Uses "*bsShowMessage"

Option Explicit


Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vSql As BPesquisa
	Set vSql = NewQuery

	vSql.Clear
	vSql.Add("SELECT 1                                                                                                      ")
	vSql.Add("  FROM SAM_BENEFICIARIO_RESPONSLEGAL                                                                          ")
	vSql.Add(" WHERE BENEFICIARIO = :BENEFICIARIO                                                                           ")
	vSql.Add("   AND HANDLE <> :HANDLE                                                                                      ")
	vSql.Add("   AND (                                                                                                      ")
	vSql.Add("           (:DATAINICIAL BETWEEN DATAINICIAL AND DATAFINAL OR :DATAFINAL BETWEEN DATAINICIAL AND DATAFINAL)   ")
	vSql.Add("        OR (DATAINICIAL BETWEEN :DATAINICIAL AND :DATAFINAL OR DATAFINAL BETWEEN :DATAINICIAL AND :DATAFINAL) ")
	vSql.Add("        OR (:DATAINICIAL >= DATAINICIAL AND DATAFINAL IS NULL))                                               ")

	vSql.ParamByName("BENEFICIARIO").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO")
	vSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	vSql.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
	vSql.ParamByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

	vSql.Active = True

	vSql.First
	If Not vSql.EOF Then
		CanContinue = False
		bsShowMessage("Não é possível gravar resgistro, pois já existe outro responsável legal em vigência no período informado!", "E")

		Set vSql = Nothing
		Exit Sub
	End If

	Set vSql = Nothing
End Sub
