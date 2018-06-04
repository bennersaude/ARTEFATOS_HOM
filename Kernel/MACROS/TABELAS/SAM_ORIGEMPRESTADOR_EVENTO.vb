'HASH: 1EE50C6B9C56AACC963CC03854863F5A
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT HANDLE FROM SAM_ORIGEMPRESTADOR_EVENTO ")
	SQL.Add(" WHERE ORIGEMPRESTADOR = " + CurrentQuery.FieldByName("ORIGEMPRESTADOR").AsString)
	SQL.Add("   AND ORIGEMEVENTO = " + CurrentQuery.FieldByName("ORIGEMEVENTO").AsString)

	SQL.Active = True

	If (Not SQL.EOF) Then
		bsShowMessage("Origem do Evento já incluída para esta Origem do Prestador.", "E")
		CanContinue = False
	End If

	Set SQL = Nothing
End Sub
