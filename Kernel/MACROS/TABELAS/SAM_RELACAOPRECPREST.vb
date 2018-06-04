'HASH: EDCF97D65031C9DD592CAA03752CB669
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Public Sub GERARARQUIVO_OnClick()

   		If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
			bsShowMessage("Uma rotina já está sendo processada no momento, verifique o monitor de processo.", "I")

			Exit Sub
		Else
			Dim qSituacao As BPesquisa
			Set qSituacao = NewQuery

			qSituacao.Add("UPDATE SAM_RELACAOPRECPREST SET SITUACAO = '4' WHERE HANDLE = :HANDLE")
            qSituacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			qSituacao.ExecSQL

			Set qSituacao = Nothing
   		End If

  		Dim sx As CSServerExec
		Set sx = NewServerExec

		Dim container As CSDContainer
		Set container = NewContainer

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		sx.Description = "PRE068 - Relação de preço do credenciado."
		sx.Process = RetornaHandleProcesso("PRE068")
		sx.SetContainer(container)
		sx.Execute

        bsShowMessage("Processo enviado ao servidor, verifique o Monitor de processos", "I")

		Set sx = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

  If CommandID = "GERARARQUIVO" Then
      GERARARQUIVO_OnClick
  End If

End Sub
