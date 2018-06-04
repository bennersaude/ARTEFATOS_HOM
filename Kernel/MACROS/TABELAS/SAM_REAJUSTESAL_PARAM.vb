'HASH: 064CFFDEB1CF8AE3E96EA559720F8112
'#Uses "*bsShowMessage"


Public Sub BOTAOCANCELAR_OnClick()
  If VisibleMode Then
	  If CurrentQuery.State = 1 Then ' O registro esta no modo browse
	  	If CurrentQuery.FieldByName("SITUACAOGERACAO").AsString <> "5" Then
	  		bsShowMessage("Situação da rotina não permite Cancelar!", "I")
	  		Exit Sub
	  	End If

	  	If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString <> "1" Then
	  		bsShowMessage("Situação da rotina não permite Cancelar. Situação processamento não está aberta.", "I")
	  		Exit Sub
	  	End If

	  	Dim vContainer As CSDContainer
	  	Set vContainer = NewContainer

	  	vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")

	  	vContainer.Insert
	  	vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  	vContainer.Field("INTERFACE").AsString = "BSBen003.ReajusteSalario_Cancelar"


	  	Dim interface As Object
	    Set interface = CreateBennerObject("BSINTERFACE.Rotinas")
	    interface.Executar(CurrentSystem,vContainer)
	    Set interface = Nothing

        WriteAudit("P", HandleOfTable("SAM_REAJUSTESAL_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Salários - Cancelar")

        CurrentQuery.Active = False
        CurrentQuery.Active = True

        Set vContainer = Nothing

	  End If
  Else
        Dim vsMensagemErro As String
   		Dim viRetorno As Long


		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen003", _
                                     "ReajusteSalario_Cancelar", _
                                     "Cancelamento Rotina de Reajuste de Salário - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_REAJUSTESAL_PARAM", _
                                     "SITUACAOGERACAO", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                      vsMensagemErro, _
                                     Null)


        If viRetorno = 0 Then
  			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      	End If

      	Set obj = Nothing
  End If

End Sub

Public Sub BOTAOGERAR_OnClick()
  If VisibleMode Then
	  If CurrentQuery.State = 1 Then ' O registro esta no modo browse
	  	If CurrentQuery.FieldByName("SITUACAOGERACAO").AsString <> "1" Then
	  		bsShowMessage("Situação da rotina não permite Gerar!", "I")
	  		Exit Sub
	  	End If

	  	Dim vContainer As CSDContainer
	  	Set vContainer = NewContainer

	  	vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")
	  	vContainer.Insert
	  	vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  	vContainer.Field("INTERFACE").AsString = "BSBen003.ReajusteSalario_Gerar"


	  	Dim interface As Object
	    Set interface = CreateBennerObject("BSINTERFACE.Rotinas")
	    interface.Executar(CurrentSystem,vContainer)
	    Set interface = Nothing

        WriteAudit("P", HandleOfTable("SAM_REAJUSTESAL_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Salários - Gerar")

        CurrentQuery.Active = False
        CurrentQuery.Active = True

        Set vContainer = Nothing


	  End If
  Else
        Dim vsMensagemErro As String
   		Dim viRetorno As Long


		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen003", _
                                     "ReajusteSalario_Gerar", _
                                     "Geração Rotina de Reajuste de Salário - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_REAJUSTESAL_PARAM", _
                                     "SITUACAOGERACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                      vsMensagemErro, _
                                     Null)


        If viRetorno = 0 Then
  			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      	End If

      	Set obj = Nothing
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If VisibleMode Then
	  If CurrentQuery.State = 1 Then ' O registro esta no modo browse
	  	If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString <> "1" Then
	  		bsShowMessage("Situação da rotina não permite Processar!", "I")
	  		Exit Sub
	  	End If

	  	Dim vContainer As CSDContainer
	  	Set vContainer = NewContainer

	  	vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")
	  	vContainer.Insert
	  	vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  	vContainer.Field("INTERFACE").AsString = "BSBen003.ReajusteSalario_Processar"


	  	Dim interface As Object
	    Set interface = CreateBennerObject("BSINTERFACE.Rotinas")
	    interface.Executar(CurrentSystem,vContainer)
	    Set interface = Nothing

        WriteAudit("P", HandleOfTable("SAM_REAJUSTESAL_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Salários - Processamento")

        CurrentQuery.Active = False
        CurrentQuery.Active = True

        Set vContainer = Nothing


	  End If
  Else
        Dim vsMensagemErro As String
   		Dim viRetorno As Long


		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen003", _
                                     "ReajusteSalario_Processar", _
                                     "Processando Rotina de Reajuste de Salário - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_REAJUSTESAL_PARAM", _
                                     "SITUACAOPROCESSAMENTO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                      vsMensagemErro, _
                                     Null)


        If viRetorno = 0 Then
  			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
    		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      	End If

      	Set obj = Nothing
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

	If CurrentQuery.FieldByName("SITUACAOGERACAO").AsString <> "1" Then
		bsShowMessage("A rotina já está gerada, não pode ser alterada!","Q")
		CanContinue = False
	End If

	If CurrentQuery.FieldByName("SITUACAOPROCESSAMENTO").AsString <> "1" Then
		bsShowMessage("A rotina já está processada, não pode ser alterada!","Q")
		CanContinue = False
	End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vCampoReajuste As Boolean

	If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
		bsShowMessage("A data final não pode ser menor que a data inicial!","Q")
		If VisibleMode Then
			DATAFINAL.SetFocus
		End If
		CanContinue = False
	End If

	If CurrentQuery.FieldByName("PERCENTUAL").AsFloat <= 0 Then
		bsShowMessage("O percentual de reajuste deve ser maior que zero!","Q")
		If VisibleMode Then
			PERCENTUAL.SetFocus
		End If
		CanContinue = False
	End If


	vCampoReajuste = CurrentQuery.FieldByName("SALARIO").AsString = "S" Or _
					 CurrentQuery.FieldByName("SALARIONORMAL").AsString = "S" Or _
					 CurrentQuery.FieldByName("VALORAPOSENTADORIA").AsString = "S" Or _
					 CurrentQuery.FieldByName("VALORAPOSENTADORIACOMPLEMENTAR").AsString = "S" Or _
					 CurrentQuery.FieldByName("DECIMOTERCEIROSALARIO").AsString = "S" Or _
					 CurrentQuery.FieldByName("OUTRASRENDAS").AsString = "S" Or _
					 CurrentQuery.FieldByName("COTAPATRONALENVIADA").AsString = "S" Or _
					 CurrentQuery.FieldByName("CONTRIBUICAOSOCIALENVIADA").AsString = "S"

	If Not vCampoReajuste Then
		bsShowMessage("Selecione um preço para reajuste!","Q")
		CanContinue = False
	End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "PROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	End If

	If CommandID = "CANCELAR" Then
		BOTAOCANCELAR_OnClick
	End If

	If CommandID = "GERAR" Then
		BOTAOGERAR_OnClick
	End If

End Sub
