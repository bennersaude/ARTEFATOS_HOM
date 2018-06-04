'HASH: D5A216F79BF1467BEFD78EB9C499C6B0
'#Uses "*bsShowMessage"
'#Uses "*ProcuraBeneficiarioAtivo"
'#Uses "*IndicarIdadeBeneficiario"
'#Uses "*VerificarBloqueioAlteracoesReapresentacao"
'#Uses "*RecordHandleOfTableInterfacePEG"
'#Uses "*RefreshNodesWithTableInterfacePEG"
'#Uses "*PermissaoAlteracao"

Option Explicit

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
	Dim vHandleBeneficiario As Long
	vHandleBeneficiario = ProcuraBeneficiarioAtivo(False,ServerDate,BENEFICIARIO.LocateText)

    If (vHandleBeneficiario <> 0) Then
        CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = vHandleBeneficiario
    End If

    ShowPopup = False
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If Not VerificarBloqueioAlteracoesReapresentacao(RecordHandleOfTableInterfacePEG("SAM_PEG")) Then
      bsShowMessage("O Beneficiário da guia não pode ser alterado porque o PEG não é de reapresentação. ", "E")
      CanContinue = False
	  Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim vSMensagem As String
	Dim vsMensagemErro As String
	Dim viRetornoPermissaoAlteracao As Integer
    Dim viHandleGuia As Long
	Dim qSql As BPesquisa
	Set qSql = NewQuery

	If Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
		viHandleGuia = RecordHandleOfTableInterfacePEG("SAM_GUIA")

		qSql.Clear
		qSql.Add("UPDATE SAM_GUIA                     ")
		qSql.Add("   SET BENEFICIARIO = :BENEFICIARIO ")
		qSql.Add(" WHERE HANDLE = :HANDLEGUIA         ")
		qSql.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
		qSql.ParamByName("HANDLEGUIA").AsInteger = viHandleGuia
		qSql.ExecSQL

		viRetornoPermissaoAlteracao = PermissaoAlteracao(0, viHandleGuia, 0, True, vSMensagem)

		If viRetornoPermissaoAlteracao = 1 Then
	       Err.Raise(vbsUserException, "", vSMensagem)
	       Exit Sub
		End If

        Dim vRetorno As Integer
        Dim vbRetorno As Boolean

		Dim vDllBSPro006 As Object
		Dim vDllSamPegDigit As Object
		Dim qGuia As BPesquisa
		Set qGuia = NewQuery

		Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuia")

		qGuia.Clear
		qGuia.Add("SELECT IDADEBENEFICIARIO,     ")
		qGuia.Add("		  DATAATENDIMENTO,       ")
		qGuia.Add("       DVCARTAO               ")
		qGuia.Add("  FROM SAM_GUIA               ")
		qGuia.Add(" WHERE HANDLE = :HANDLE       ")
		qGuia.ParamByName("HANDLE").AsInteger = viHandleGuia
		qGuia.Active = True

		vRetorno = vDllBSPro006.IndicarIdadeBeneficiario(CurrentSystem, _
											  		     qGuia.FieldByName("DATAATENDIMENTO").AsDateTime, _
											  			 viHandleGuia, _
											  			 CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, _
											  			 qGuia.FieldByName("IDADEBENEFICIARIO").AsInteger, _
											  			 True)

		vbRetorno = vDllBSPro006.IndicarDvCartao(CurrentSystem, _
                                     			 qGuia.FieldByName("DATAATENDIMENTO").AsDateTime, _
                                     		     viHandleGuia, _
                                     			 CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, _
                                     			 qGuia.FieldByName("DVCARTAO").AsString, _
                                     			 True, _
                                     			 "")


		Set vDllBSPro006 = Nothing

		Set vDllSamPegDigit = CreateBennerObject("SAMPEGDIGIT.Rotinas")

		qGuia.Clear
		qGuia.Add("SELECT IDADEBENEFICIARIO,     ")
		qGuia.Add("		  DATAATENDIMENTO,       ")
		qGuia.Add("       HORAATENDIMENTO,       ")
		qGuia.Add("       DVCARTAO,              ")
		qGuia.Add("       EXECUTOR,              ")
		qGuia.Add("       RECEBEDOR,             ")
		qGuia.Add("       MODELOGUIA,            ")
		qGuia.Add("       FINALIDADEATENDIMENTO, ")
		qGuia.Add("       TABREGIMEPGTO          ")
		qGuia.Add("  FROM SAM_GUIA               ")
		qGuia.Add(" WHERE HANDLE = :HANDLE       ")
		qGuia.ParamByName("HANDLE").AsInteger = viHandleGuia
		qGuia.Active = True

		vDllSamPegDigit.CopiarAlteracoesParaEventos( CurrentSystem, _
                                      				 qGuia.FieldByName("EXECUTOR").AsInteger, _
                                      				 qGuia.FieldByName("RECEBEDOR").AsInteger, _
                                      				 CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, _
                                      				 qGuia.FieldByName("IDADEBENEFICIARIO").AsInteger, _
                                      				 qGuia.FieldByName("DVCARTAO").AsString, _
                                      			     viHandleGuia, _
                                      				 qGuia.FieldByName("MODELOGUIA").AsInteger, _
                                      				 qGuia.FieldByName("DATAATENDIMENTO").AsDateTime, _
                                      				 qGuia.FieldByName("HORAATENDIMENTO").AsDateTime, _
                                      				 qGuia.FieldByName("FINALIDADEATENDIMENTO").AsInteger, _
                                      			 	 qGuia.FieldByName("TABREGIMEPGTO").AsInteger)
		Set vDllSamPegDigit = Nothing
		qGuia.Active = False
		Set qGuia = Nothing

		If VisibleMode Then
		   Dim Reprocessar As Object

		   Set Reprocessar = CreateBennerObject("BSPro000.Rotinas")
		   Reprocessar.VerificarEventosGuia(CurrentSystem, viHandleGuia)
		   Set Reprocessar = Nothing
	  	Else
	    	Dim Obj As Object
	    	Dim viRet As Long
	    	Dim vcContainer As CSDContainer
	   		Set vcContainer = NewContainer
	   		vcContainer.AddFields("HANDLE:INTEGER")

		    vcContainer.Insert
		    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

		    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		    viRet = Obj.ExecucaoImediata(CurrentSystem, _
		                               	 "BSPRO000", _
		                               	 "VerificarEventosGuia", _
		                               	 "Reprocessamento de GUIA", _
		                               	 viHandleGuia, _
		                               	 "SAM_GUIA", _
		                               	 "SITUACAOPROCESSAMENTO", _
		                               	 "", _
		                               	 "", _
		                               	 "P", _
		                               	 True, _
		                               	 vsMensagemErro, _
		                               	 vcContainer)

			If viRet = 0 Then
			 	bsShowMessage("Processo enviado para execução no servidor!", "I")
			Else
		     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		   	End If
	  	End If

		bsShowMessage("Alteração Concluída","I")

		Set qSql = Nothing
		RefreshNodesWithTableInterfacePEG("SAM_GUIA")
	End If
End Sub
