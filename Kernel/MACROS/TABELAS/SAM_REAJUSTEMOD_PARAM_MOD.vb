'HASH: 745FE3FA4436395EF2E61946985AC61F
'TABELA SAM_REAJUSTEMOD_PARAM_MOD

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If Solver(CurrentQuery.FieldByName("REAJUSTEPARAM").AsInteger,"SAM_REAJUSTEMOD_PARAM","SITUACAO") <> "1" Then
		bsShowMessage("Situação da rotina não permite excluir registro!", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "GERAR" Then
		Dim vContainer As CSDContainer
		Set vContainer = NewContainer

		If SessionVar("CONTAINER") = "" Then
  			bsShowMessage("A geração somente permitida, após consultar a Interface de pesquisa.", "E")
  			CanContinue = False
		Else


			vContainer.SetXML(SessionVar("CONTAINER"),True, True, True)

	        Dim vsMensagemErro As String
	   		Dim viRetorno As Long

			Dim obj As Object
	        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
	                                     "BSBen003", _
	                                     "RotinaReajuste_Gerar", _
	                                     "Geração da Rotina de Reajuste de Contrato - Rotina: " + CStr(vContainer.Field("HANDLE").AsInteger), _
	                                     vContainer.Field("HANDLE").AsInteger, _
	                                     "SAM_REAJUSTEMOD_PARAM", _
	                                     "SITUACAOGERAR", _
	                                     "", _
	                                     "", _
	                                     "P", _
	                                     True, _
	                                     vsMensagemErro, _
	                                     vContainer)


	        If viRetorno = 0 Then
	  			bsShowMessage("Processo enviado para execução no servidor!", "I")
		 	Else
	    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	      	End If

	      	Set obj = Nothing
	    End If

	    Set vContainer = Nothing


	End If
End Sub
