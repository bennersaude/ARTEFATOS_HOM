'HASH: 209CE5A7E68734468BB338E8F4495CFA
'#Uses  "*bsShowMessage"

Public Sub TABLE_AfterPost()
		Dim vContainer As CSDContainer
		Set vContainer = NewContainer



		vContainer.GetFieldsFromQuery(CurrentQuery.TQuery)
		vContainer.LoadAllFromQuery(CurrentQuery.TQuery)


        Dim vsMensagemErro As String
   		Dim viRetorno As Long

		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen014", _
                                     "EncerramentoSuspensao_Processar", _
                                     "Processamento da Rotina de Encerramento de Suspensão", _
                                     0, _
                                     "", _
                                     "", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     vContainer)


        If viRetorno = 0 Then
  			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      	End If

      	Set obj = Nothing

	    Set vContainer = Nothing


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("BAIXARSUSPENSAOFATURAPAGA").AsInteger = 2 Then
		If CurrentQuery.FieldByName("QTDFATURASATRASO").AsInteger = 0 Then
			CancelDescription = "''Qtd de fatura atraso'' deve ser maior que zero"
			CanContinue = False
		End If
	End If
End Sub
