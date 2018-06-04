'HASH: 99EE6326CA43D7A91BAA8DC295CC675E
'#Uses "*bsShowMessage"
Dim qOperacao As Object

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Set qOperacao = NewQuery

	qOperacao.Clear
	qOperacao.Add("SELECT HANDLE                  ")
	qOperacao.Add("   FROM SIS_OPERACAO           ")
	qOperacao.Add(" WHERE CODIGO = :CODIGO        ")

	qOperacao.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
	qOperacao.Active = True

	If qOperacao.FieldByName("HANDLE").AsInteger > 0 Then
		MsgBox("Esse código já existe, por favor digite outro!")
		CODIGO.SetFocus
		CanContinue = False
	ElseIf CurrentQuery.FieldByName("CONSIDERACONSULTASALD").AsString = "N" Then
    	bsShowMessage("Esta operação não será considerada nas consultas de saldo devedor da conta PF", "I")
	End If

  	Set qOperacao= Nothing
End Sub

