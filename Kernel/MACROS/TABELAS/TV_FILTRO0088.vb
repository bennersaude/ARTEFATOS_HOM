'HASH: 00424D2F38D671AFEF7BF82A9D4D9DAB
 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim qAux As Object

	Set qAux = NewQuery
	qAux.Clear
	qAux.Active = False
	qAux.Add("SELECT CODIGO")
	qAux.Add("  FROM SIS_TIPOFATURAMENTO")
	qAux.Add(" WHERE HANDLE = :HANDLE")
	qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
	qAux.Active = True

	If qAux.FieldByName("CODIGO").AsInteger <> 610 Then
		bsShowMessage("Tipo de Faturamento somente deve ser '610.Pagamento de INSS'", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
