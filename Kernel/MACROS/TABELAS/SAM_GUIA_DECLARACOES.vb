'HASH: 6551B3DBFB6383ACDADCB068B52A4231
'#Uses "*bsShowMessage"
'#Uses "*VerificaPermissaoEdicaoTriagem"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	CanContinue = VerificarPermissaoUsuarioPegTriado(True)
	RecordReadOnly = Not CanContinue
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If Not VerificarPermissaoUsuarioPegTriado(True) Then
 		CanContinue = False
 		Exit Sub
 	End If

	Dim qVerificaQtdeDecalarcoes As Object
	Set qVerificaQtdeDecalarcoes = NewQuery

	If CurrentQuery.State = 3 Then
		qVerificaQtdeDecalarcoes.Clear
		qVerificaQtdeDecalarcoes.Add("SELECT COUNT(1) QTD")
		qVerificaQtdeDecalarcoes.Add("  FROM SAM_GUIA_DECLARACOES ")
		qVerificaQtdeDecalarcoes.Add(" WHERE GUIA = :GUIA         ")
		qVerificaQtdeDecalarcoes.ParamByName("GUIA").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
		qVerificaQtdeDecalarcoes.Active = True

		If qVerificaQtdeDecalarcoes.FieldByName("QTD").AsInteger >= 8 Then
			bsShowMessage("São permitidas apenas 8 declarações por guia!", "I")
			CanContinue = False
			CancelDescription = "São permitidas apenas 8 declarações por guia!"
		End If
	End If

	Set qVerificaQtdeDecalarcoes = Nothing
End Sub
