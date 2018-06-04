'HASH: E042531AB401E09256729254B1F76B08
'Macro: SAM_ROTINAIMP
'Juliano 16/10/00
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELA_OnClick()
  Dim IMPORTA As Object
  Dim viRetornoMensagem As Long
  Dim vsMensagemErro As String


  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("TABTIPOIMPORTACAO").AsString = "3" Then
    'Set IMPORTA = CreateBennerObject("BSBEN015.ROTINAS")
    'IMPORTA.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    If VisibleMode Then
      Set IMPORTA = CreateBennerObject("BSINTERFACE0015.RotinasImportacaoBenef")
      viRetornoMensagem = IMPORTA.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Else
      Set IMPORTA = CreateBennerObject("BSBEN015.ImportarCancelar")
      viRetornoMensagem = IMPORTA.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemErro, 0 )
    End If

    If viRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If


  Else
    'Set IMPORTA = CreateBennerObject("BSBEN005.ROTINAS")
    'IMPORTA.Cancelar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)

    Dim qVerificaFilialConfirmada As Object
    Set qVerificaFilialConfirmada = NewQuery

    qVerificaFilialConfirmada.Active = False
    qVerificaFilialConfirmada.Clear
    qVerificaFilialConfirmada.Add("SELECT F.DATAHORACONFIRMA       ")
    qVerificaFilialConfirmada.Add("  FROM SAM_ROTINAIMP_FILIAL F,  ")
    qVerificaFilialConfirmada.Add("       SAM_ROTINAIMP R          ")
    qVerificaFilialConfirmada.Add(" WHERE R.Handle = F.ROTINAIMP   ")
    qVerificaFilialConfirmada.Add("   And F.DATAHORACONFIRMA Is Not Null ")
    qVerificaFilialConfirmada.Add("   And F.ROTINAIMP = :HANDLEROTINA    ")
    qVerificaFilialConfirmada.ParamByName("HANDLEROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qVerificaFilialConfirmada.Active = True

    If Not qVerificaFilialConfirmada.EOF Then
      bsShowMessage("Já foi Confirmado","I")
      Set qVerificaFilialConfirmada = Nothing
      Exit Sub
    ElseIf Not CurrentQuery.FieldByName("DATAHORAEXCLUSAO").IsNull Then
      bsShowMessage("A Rotina não pode ser Cancelada se já estiver excluída !","I")
      Set qVerificaFilialConfirmada = Nothing
      Exit Sub
    ElseIf Not CurrentQuery.FieldByName("DATAHORACANCELA").IsNull Then
      bsShowMessage("Já foi Cancelado!","I")
      Set qVerificaFilialConfirmada = Nothing
      Exit Sub
    End If


    Set qVerificaFilialConfirmada = Nothing

    If bsShowMessage("A Rotina será cancelada, Deseja Continuar?","Q") Then

      If VisibleMode Then
        Set IMPORTA = CreateBennerObject("BSINTERFACE0025.RotinasImportacaoBenef")
        viRetornoMensagem = IMPORTA.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      Else
        Set IMPORTA = CreateBennerObject("BSBEN005.RotinaImportar_Cancelar")
        viRetornoMensagem = IMPORTA.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMensagemErro, 0 )
        bsShowMessage("Cancelamento executado com sucesso.", "I")
      End If

      If viRetornoMensagem = 1 Then
        bsShowMessage("Ocorreu erro no processo","I")
      End If
    End If


  End If

  Set IMPORTA = Nothing

  WriteAudit("C", HandleOfTable("SAM_ROTINAIMP"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
  	  "Rotina de Atualização de Cadastro - Cancelar Importação de Beneficiários")

  If VisibleMode Then
	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOEXCLUIR_OnClick()
  If Not CurrentQuery.FieldByName("DATAHORAEXCLUSAO").IsNull Then
    bsShowMessage("A Rotina já foi Excluída !", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAHORAPROCESSO").IsNull Then
    Dim VERIFICA As Object
    Set VERIFICA = NewQuery

    VERIFICA.Active = False

    VERIFICA.Clear

    VERIFICA.Add("SELECT DATAHORACONFIRMA ")
    VERIFICA.Add("  FROM SAM_ROTINAIMP_FILIAL")
    VERIFICA.Add(" WHERE ROTINAIMP = :HANDLE")
    VERIFICA.Add("   AND DATAHORACONFIRMA IS NOT NULL")

    VERIFICA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    VERIFICA.Active = True

    If VERIFICA.EOF Then
      bsShowMessage("A Rotina precisa ser confirmada ou cancelada antes de ser excluída !", "I")
      Exit Sub
    End If

    Set VERIFICA = Nothing
  End If

  If bsShowMessage("Confirma a Exclusão da Rotina ?", "Q") = vbYes Then
	Dim EXCLUI As Object
	Set EXCLUI = NewQuery

	If Not InTransaction Then StartTransaction

	EXCLUI.Add("DELETE SAM_ROTINAIMP_BENEF_HOMO")
	EXCLUI.Add("  FROM SAM_ROTINAIMP_FILIAL FL,")
	EXCLUI.Add("       SAM_ROTINAIMP_FAM F,")
	EXCLUI.Add("       SAM_ROTINAIMP_BENEF B,")
	EXCLUI.Add("       SAM_ROTINAIMP_BENEF_HOMO H")
	EXCLUI.Add(" WHERE ROTINAIMP = :HANDLE")
	EXCLUI.Add("   AND F.ROTINAIMPFILIAL = FL.HANDLE")
	EXCLUI.Add("   AND B.IMPFAM = F.HANDLE")
	EXCLUI.Add("   AND H.IMPORTABENEF = B.HANDLE")

	EXCLUI.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	EXCLUI.ExecSQL

	EXCLUI.Add("DELETE SAM_ROTINAIMP_BENEF")
	EXCLUI.Add("  FROM SAM_ROTINAIMP_FILIAL FL,")
	EXCLUI.Add("       SAM_ROTINAIMP_FAM F,")
	EXCLUI.Add("       SAM_ROTINAIMP_BENEF B")
	EXCLUI.Add(" WHERE ROTINAIMP = :HANDLE")
	EXCLUI.Add("   AND F.ROTINAIMPFILIAL = FL.HANDLE")
	EXCLUI.Add("   AND B.IMPFAM = F.HANDLE")

	EXCLUI.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	EXCLUI.ExecSQL

	EXCLUI.Add("DELETE SAM_ROTINAIMP_FAM ")
	EXCLUI.Add("  FROM SAM_ROTINAIMP_FILIAL FL,")
	EXCLUI.Add("       SAM_ROTINAIMP_FAM F")
	EXCLUI.Add(" WHERE ROTINAIMP = :HANDLE")
	EXCLUI.Add("   AND F.ROTINAIMPFILIAL = FL.HANDLE")

	EXCLUI.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	EXCLUI.ExecSQL

	EXCLUI.Add("DELETE SAM_ROTINAIMP_FILIAL")
	EXCLUI.Add(" WHERE ROTINAIMP = :HANDLE")

	EXCLUI.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	EXCLUI.ExecSQL

	Set EXCLUI = Nothing
	Dim ATUALIZA As Object
	Set ATUALIZA = NewQuery

	ATUALIZA.Add("UPDATE SAM_ROTINAIMP")
	ATUALIZA.Add("   SET USUARIOEXCLUSAO  = :USUARIO,")
	ATUALIZA.Add("       DATAHORAEXCLUSAO = :DATAHORA,")
	ATUALIZA.Add("       USUARIOCANCELA   = NULL,")
	ATUALIZA.Add("       DATAHORACANCELA  = NULL,")
	ATUALIZA.Add("       USUARIOPROCESSO  = NULL,")
	ATUALIZA.Add("       DATAHORAPROCESSO = NULL")
	ATUALIZA.Add(" WHERE HANDLE = :HANDLE")

	ATUALIZA.ParamByName("USUARIO").Value = CurrentUser
	ATUALIZA.ParamByName("DATAHORA").Value = ServerDate
	ATUALIZA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	ATUALIZA.ExecSQL

	If InTransaction Then Commit

	Set ATUALIZA = Nothing

	CurrentQuery.Active = False
	CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOIMPORTA_OnClick()
  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  If Not (CurrentQuery.FieldByName("DATAHORAPROCESSO").IsNull) Then
    bsShowMessage(" Já foi Processado","I")
    Exit Sub
  End If

  Dim IMPORTA As Object
  Dim vsRetornoMensagem As Long
  Dim vsMensagemErro As String
  Dim viRetorno As Long
  Dim dllBSServerExec As Object

  If CurrentQuery.FieldByName("TABTIPOIMPORTACAO").AsString = "3" Then
  	If VisibleMode Then
	    Set IMPORTA = CreateBennerObject("BSINTERFACE0015.RotinasImportacaoBenef")
	    vsRetornoMensagem = IMPORTA.Verificar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	    If vsRetornoMensagem = 1 Then
	      bsShowMessage("Ocorreu erro no processo","I")
	    End If
    Else
        Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")

        viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                     "BSBEN015", _
                                                     "ImportarBeneficiarios", _
                                                     "Importação de beneficiários", _
                                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                     "SAM_ROTINAIMP", _
                                                     "SITUACAOROTINAIMP", _
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
    End If
  Else
    Dim VERIFICA As Object
    Set VERIFICA = NewQuery

  	VERIFICA.Active = False
  	VERIFICA.Clear
  	VERIFICA.Add("SELECT TABTIPOLEIAUTEIMP FROM SAM_PARAMETROSBENEFICIARIO")
  	VERIFICA.Active = True

  	If VERIFICA.FieldByName("TABTIPOLEIAUTEIMP").AsInteger = 2 Then
	  'Set IMPORTA = CreateBennerObject("BSBEN005.ROTINAS")
	  'IMPORTA.VERIFICAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

      If VisibleMode Then
        Set IMPORTA = CreateBennerObject("BSINTERFACE0025.RotinasImportacaoBenef")
	    viRetorno = IMPORTA.Verificar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	    If viRetorno = 1 Then
	      bsShowMessage("Ocorreu erro no processo","I")
	    End If

	  Else
        Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")

        viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                     "BSBEN005", _
                                                     "RotinaImportar_Verificar", _
                                                     "Importação de beneficiários", _
                                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                     "SAM_ROTINAIMP", _
                                                     "SITUACAOROTINAIMP", _
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


	  End If

  	Else
      If VisibleMode Then
        Set IMPORTA = CreateBennerObject("BSINTERFACE0034.RotinasImportacaoBenef")
	    viRetorno = IMPORTA.Verificar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	    If viRetorno = 1 Then
	      bsShowMessage("Ocorreu erro no processo","I")
	    End If

	  Else
        Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")

        viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                     "BSBEN004", _
                                                     "RotinaImportar_Verificar", _
                                                     "Importação de beneficiários", _
                                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                     "SAM_ROTINAIMP", _
                                                     "SITUACAOROTINAIMP", _
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

	  End If
  	End If

	Set IMPORTA = Nothing
  	Set VERIFICA = Nothing
  	Set dllBSServerExec = Nothing

  	WriteAudit("I", HandleOfTable("SAM_ROTINAIMP"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
  		"Rotina de Atualização de Cadastro - Importação de Beneficiários")

	If VisibleMode Then
	  SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
	End If
  End If

  Set IMPORTA = Nothing
End Sub

Public Sub BOTAOOCORRENCIAS_OnClick()
  Dim exportarOcorrencias As CSBusinessComponent

  Set exportarOcorrencias = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.RotinaImportacao.SamRotinaImpBLL, Benner.Saude.Beneficiarios.Business")
  exportarOcorrencias.AddParameter(pdtInteger,CurrentQuery.FieldByName("HANDLE").AsInteger)

  exportarOcorrencias.Execute("ExportarOcorrenciasImportacaoBeneficiarios")

  bsShowMessage("Processamento enviado para execução no servidor.", "I")

  Set exportarOcorrencias = Nothing
  
End Sub

Public Sub BOTAORETORNO_OnClick()
  Dim vsMensagemErro As String
  
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim IMPORTA As Object
  Dim vsRetornoMensagem As Long

  If VisibleMode Then
    Set IMPORTA = CreateBennerObject("BSINTERFACE0015.RotinasImportacaoBenef")
    vsRetornoMensagem = IMPORTA.RETORNO(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    If vsRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If


  Else
      Set IMPORTA = CreateBennerObject("BSServerExec.ProcessosServidor")
      vsRetornoMensagem = IMPORTA.ExecucaoImediata(CurrentSystem, _
                                                   "BSBEN015", _
                                                   "ImportarRetorno", _
                                                   "Retorno dos beneficiários cadastrados", _
                                                   CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                   "SAM_ROTINAIMP", _
                                                   "SITUACAOEXPORTAR", _
                                                   "", _
                                                   "", _
                                                   "P", _
                                                   False, _
                                                   vsMensagemErro, _
                                                   Null)

        If vsRetornoMensagem = 0 Then
           bsShowMessage("Processo enviado para execução no servidor!", "I")
        Else
           bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
        End If

 End If


  Set IMPORTA = Nothing

  If VisibleMode Then
	SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If


End Sub

Public Sub TABLE_AfterScroll()
	ARQUIVO.Visible = False
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAOCANCELA"
      BOTAOCANCELA_OnClick
    Case "BOTAOEXCLUIR"
      BOTAOEXCLUIR_OnClick
    Case "BOTAOIMPORTA"
        BOTAOIMPORTA_OnClick
    Case "BOTAORETORNO"
      BOTAORETORNO_OnClick
    Case "BOTAOOCORRENCIAS"
      BOTAOOCORRENCIAS_OnClick
  End Select
End Sub
