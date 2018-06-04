'HASH: 65337D2E11AC4649EC71DC4FCAC49231
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object

	Set SQL = NewQuery
    SQL.Clear
	SQL.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE= " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
	SQL.Active = True



    If (SQL.FieldByName("CODIGO").AsInteger <> 640) Then
	  bsShowMessage("Tipo de Faturamento deve ser '640 - Recolhimento de ISS' !", "E")
	  Set SQL = Nothing
	  CanContinue = False
	  Exit Sub
	End If

	Set SQL = Nothing
End Sub
