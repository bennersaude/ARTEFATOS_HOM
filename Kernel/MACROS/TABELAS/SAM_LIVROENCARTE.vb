'HASH: C4CFA5F7C5C864A027DA42DCF77F7FC5
'SAM_LIVROENCARTE
'Matheus - sms 4572
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
	Dim SQL As Object
	Dim Msg As String
	Dim Result As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "I")
		CanContinue = False
		Exit Sub
	End If

	If bsShowMessage("Esta operação irá excluir todos os dados deste encarte." + (Chr(13)) + _
		   "Deseja continuar?", "Q") = vbYes Then

		Set SQL = NewQuery

		SQL.Add("SELECT ROTINA FROM SAM_LIVRO_ROTINAEXPORTACAO WHERE LIVRO = :LIVRO AND LIVROENCARTE = :LIVROENCARTE")

		SQL.ParamByName("LIVRO").Value = CurrentQuery.FieldByName("LIVRO").Value
		SQL.ParamByName("LIVROENCARTE").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		If Not SQL.EOF Then
			bsShowMessage("Encarte não pode ser excluído, pois o mesmo está na rotina de exportação " + _
				SQL.FieldByName("ROTINA").AsString, "I")
			Set SQL = Nothing
			Exit Sub
		End If

		SQL.Active = False

		SQL.Clear

		SQL.Add("SELECT ROTINA FROM SAM_LIVRO_ROTINAEMISSAO WHERE LIVRO = :LIVRO AND LIVROENCARTE = :LIVROENCARTE")

		SQL.ParamByName("LIVRO").Value = CurrentQuery.FieldByName("LIVRO").Value
		SQL.ParamByName("LIVROENCARTE").Value = CurrentQuery.FieldByName("HANDLE").Value
		SQL.Active = True

		If Not SQL.EOF Then
			bsShowMessage("Encarte não pode ser excluído, pois o mesmo está na rotina de emissão " + _
				SQL.FieldByName("ROTINA").AsString, "I")
			Set SQL = Nothing
			Exit Sub
		End If

		'SMS 90283 - Ricardo Rocha - Adequacao WEB
		Set Obj = CreateBennerObject("BSPRE001.Rotinas_CancelarEncarte")
		Result = Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("LIVRO").AsInteger)

		If Result <> "" Then
			bsShowMessage(Result, "I")
		End If

	Set SQL = Nothing

  	End If

  	RefreshNodesWithTable("SAM_LIVROENCARTE")

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
	End Select
End Sub
