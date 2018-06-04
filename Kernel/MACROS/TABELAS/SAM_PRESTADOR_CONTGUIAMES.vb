'HASH: 2354331ACB92CD0FD74DAA4D87226120
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim qMunicipio As Object
	Set qMunicipio = NewQuery

	qMunicipio.Active = False
	qMunicipio.Add("SELECT MUNICIPIOPAGAMENTO FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
	qMunicipio.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
	qMunicipio.Active = True

	If checkPermissao (CurrentSystem, CurrentUser, "M", qMunicipio.FieldByName("MUNICIPIOPAGAMENTO").AsInteger, "E") = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode excluir", "E")
		CanContinue = False
		Set qMunicipio = Nothing
		Exit Sub
	End If

	Set qMunicipio = Nothing

	If CurrentUser <> CurrentQuery.FieldByName("USUARIO").AsInteger Then
		CanContinue = False
		bsShowMessage("Operação cancelada. Usuário diferente", "E")
		Exit Sub
	End If
End Sub
