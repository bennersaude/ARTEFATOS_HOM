'HASH: F8D54EFF30866B96E6D9DA3D71BEC299
'Macro: SAM_REAJUSTEMOD_PARAM

'#Uses "*SAM_REAJUSTEMOD_Excluir"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
	If VisibleMode Then
	  If CurrentQuery.State = 1 Then ' O registro esta no modo browse

	  	If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
	  		bsShowMessage("Situação da rotina não permite cancelamento!", "I")
	  		Exit Sub
	  	End If

	  	 Dim vContainer As CSDContainer
	  	 Set vContainer = NewContainer

	  	 vContainer.AddFields ("HANDLE:INTEGER;INTERFACE:STRING")

	  	 vContainer.Insert
		 vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		 vContainer.Field("INTERFACE").AsString = "BSBen003.RotinaReajuste_Cancelar"


	      Dim interface As Object
	      Set interface = CreateBennerObject("BSINTERFACE.Rotinas")
	      interface.Executar(CurrentSystem,vContainer)

	      WriteAudit("P", HandleOfTable("SAM_REAJUSTEMOD_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Contratos - Cancelamento")

	      CurrentQuery.Active  = False
	      CurrentQuery.Active  = True

	      Set vContainer = Nothing
	      Set interface  = Nothing

	  End If
	Else

        Dim vsMensagemErro As String
   		Dim viRetorno As Long

		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen003", _
                                     "RotinaReajuste_Cancelar", _
                                     "Cancelando Rotina de Reajuste de Contrato - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_REAJUSTEMOD_PARAM", _
                                     "SITUACAO", _
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
  If CurrentQuery.State = 1 Then ' O registro esta no modo browse
    If CurrentQuery.FieldByName("SITUACAO").AsInteger = 1 Then
      Dim interface As Object
      Set interface = CreateBennerObject("BSBen003.Rotinas")
      interface.Gerar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      Set interface = Nothing

      WriteAudit("G", HandleOfTable("SAM_REAJUSTEMOD_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Contratos - Geração")
      RefreshNodesWithTable("SAM_REAJUSTEMOD_PARAM")
    Else
		bsShowMessage("A geração somente é permitida com a situação do Processamento em aberto.", "E")
    End If
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If VisibleMode Then
	  If CurrentQuery.State = 1 Then ' O registro esta no modo browse
	  	If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
	  		bsShowMessage("Situação da rotina não permite Processar!", "I")
	  		Exit Sub
	  	End If

	  	 Dim vContainer As CSDContainer
	  	 Set vContainer = NewContainer

	  	 vContainer.AddFields ("HANDLE:INTEGER;INTERFACE:STRING")

	  	 vContainer.Insert
		 vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		 vContainer.Field("INTERFACE").AsString = "BSBen003.RotinaReajuste_Processar"

	  	Dim interface As Object
	    Set interface = CreateBennerObject("BSINTERFACE.Rotinas")
	    interface.Executar(CurrentSystem, vContainer)

        WriteAudit("P", HandleOfTable("SAM_REAJUSTEMOD_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Contratos - Processamento")

	    CurrentQuery.Active  = False
	    CurrentQuery.Active  = True

	    Set vContainer = Nothing


	  End If
  Else
        Dim vsMensagemErro As String
   		Dim viRetorno As Long

		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen003", _
                                     "RotinaReajuste_Processar", _
                                     "Processando Rotina de Reajuste de Contrato - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_REAJUSTEMOD_PARAM", _
                                     "SITUACAO", _
                                     "SITUACAOGERAR", _
                                     "A geração dos contratos ainda não foi processada.", _
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

Public Sub BOTAORECALCULAR_OnClick()

	If VisibleMode Then
		If CurrentQuery.State = 1 Then ' O registro esta no modo browse
		    If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
		  		bsShowMessage("Situação da rotina não permite Recalcular!", "I")
		  		Exit Sub
	  		End If

			 Dim vContainer As CSDContainer
			 Set vContainer = NewContainer

			 vContainer.AddFields ("HANDLE:INTEGER;INTERFACE:STRING")

			 vContainer.Insert
			 vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			 vContainer.Field("INTERFACE").AsString = "BSBen003.RotinaReajuste_Recalcular"

			Dim interface As Object
			Set interface = CreateBennerObject("BSINTERFACE.Rotinas")
			interface.Executar(CurrentSystem,vContainer)

			WriteAudit("G", HandleOfTable("SAM_REAJUSTEMOD_PARAM"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Reajuste de Contratos - Recalcular")

            CurrentQuery.Active  = False
	        CurrentQuery.Active  = True

	        Set vContainer = Nothing
			Set interface = Nothing

		End If

    Else

        Dim vsMensagemErro As String
   		Dim viRetorno As Long

		Dim obj As Object
        Set obj = CreateBennerObject("BSServerExec.ProcessosServidor")
        viRetorno = obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen003", _
                                     "RotinaReajuste_Recalcular", _
                                     "Recalculando Rotina de Reajuste de Contrato - Rotina: " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_REAJUSTEMOD_PARAM", _
                                     "SITUACAORECALCULO", _
                                     "SITUACAOGERAR", _
                                     "A geração dos contratos ainda não foi processada.", _
                                     "P", _
                                     True, _
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

Public Sub INDICEDEREAJUSTE_OnExit()
	If CurrentQuery.State = 2 Then
  		CurrentQuery.FieldByName("INDICEREAJUSTEFAMILIA").AsFloat = CurrentQuery.FieldByName("INDICEDEREAJUSTE").AsFloat
  	End If
End Sub

Public Sub TABLE_AfterScroll()
	SessionVar("REAJUSTEPARAM") = CurrentQuery.FieldByName("HANDLE").AsString
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	SAM_REAJUSTEMOD_Excluir 1, CurrentQuery.FieldByName("HANDLE").AsInteger

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("DATADOPROCESSO").IsNull Then
    bsShowMessage("Reajuste já processado. Operação não permitida.","I")
    CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("TABREFERENCIAANIVERSARIO").AsInteger = 1 Then 'contrato
    CurrentQuery.FieldByName("REAJUSTACONTRATOMOD").Value = "N"
    CurrentQuery.FieldByName("REAJUSTAFAMILIAMOD").Value = "N"
    CurrentQuery.FieldByName("TABREFERENCIAMOD").Value = 0
    CurrentQuery.FieldByName("DATAINICIALADESAO").Clear
    CurrentQuery.FieldByName("DATAFINALADESAO").Clear
    CurrentQuery.FieldByName("DATAINICIALCOMPETENCIA").Clear
    CurrentQuery.FieldByName("DATAFINALCOMPETENCIA").Clear
  Else
    If CurrentQuery.FieldByName("TABREFERENCIAANIVERSARIO").AsInteger = 2 Then 'modulo
      CurrentQuery.FieldByName("DATAINICIAL").Clear
      CurrentQuery.FieldByName("DATAFINAL").Clear
      CurrentQuery.FieldByName("PRORATAMODULO").Clear
      CurrentQuery.FieldByName("REAJUSTACONTRATO").Value = "N"
      CurrentQuery.FieldByName("REAJUSTAFAMILIA").Value = "N"
      CurrentQuery.FieldByName("PRORATAMODULO").Value = "N"
      If CurrentQuery.FieldByName("TABREFERENCIAMOD").AsInteger = 1 Then
        CurrentQuery.FieldByName("DATAINICIALCOMPETENCIA").Clear
        CurrentQuery.FieldByName("DATAFINALCOMPETENCIA").Clear
      Else
        CurrentQuery.FieldByName("DATAINICIALADESAO").Clear
        CurrentQuery.FieldByName("DATAFINALADESAO").Clear
      End If
    Else 'familia
      If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) And _
          (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
        bsShowMessage("A data final para o filtro por aniversário deve ser maior que a data inicial.","I")
        CanContinue = False
        Exit Sub
      End If

      CurrentQuery.FieldByName("REAJUSTAFAMILIA").Value = "S"
      CurrentQuery.FieldByName("REAJUSTACONTRATO").Value = "N"
      CurrentQuery.FieldByName("REAJUSTACONTRATOMOD").Value = "N"
      CurrentQuery.FieldByName("REAJUSTAFAMILIAMOD").Value = "N"
      CurrentQuery.FieldByName("TABREFERENCIAMOD").Value = 0
      CurrentQuery.FieldByName("DATAINICIALADESAO").Clear
      CurrentQuery.FieldByName("DATAFINALADESAO").Clear
      CurrentQuery.FieldByName("DATAINICIALCOMPETENCIA").Clear
      CurrentQuery.FieldByName("DATAFINALCOMPETENCIA").Clear
    End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	If CommandID = "PROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	End If

	If CommandID = "CANCELAR" Then
		BOTAOCANCELAR_OnClick
	End If

	If CommandID = "RECALCULAR" Then
		BOTAORECALCULAR_OnClick
	End If

End Sub

Public Sub TABREFERENCIAANIVERSARIO_OnChanging(AllowChange As Boolean)
  If Not CurrentQuery.FieldByName("DATADOPROCESSO").IsNull Then
    AllowChange = False
     bsShowMessage("Reajuste já processado. Operação não permitida.","I")
  End If
End Sub




