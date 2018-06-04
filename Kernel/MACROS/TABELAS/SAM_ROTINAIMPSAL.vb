'HASH: AC97A089D8A1EDB52E1D888087CAF6BC
'Macro da tabela: SAM_ROTINAIMPSAL
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOIMPORTAR_OnClick()

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim InterfDll As Object
  Dim vsRetornoMensagem As Long
  Dim vsMensagemErro As String
  Dim viRetorno As Long
  Dim dllBSServerExec As Object

  If VisibleMode Then
    Set InterfDll = CreateBennerObject("BSINTERFACE0060.RotinasImportacaoSalario")
    vsRetornoMensagem = InterfDll.Importar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    If vsRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If
  Else
    Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")

    viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                 "BSBEN033", _
                                                 "ImportacaoSalario", _
                                                 "Importação de salários", _
                                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                 "SAM_ROTINAIMPSAL", _
                                                 "SITUACAOIMPORTAR", _
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

	Set InterfDll = Nothing
  	Set dllBSServerExec = Nothing

  	WriteAudit("I", HandleOfTable("SAM_ROTINAIMPSAL"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
  		"Rotina de Atualização de Cadastro - Importação de salários")

	If VisibleMode Then
	  SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
	End If
  End If

End Sub

Public Sub BOTAOCANCELAR_OnClick()

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim InterfDll As Object
  Dim vsRetornoMensagem As Long
  Dim vsMensagemErro As String
  Dim viRetorno As Long
  Dim dllBSServerExec As Object
  Dim SQL As Object

  If VisibleMode Then
    Set InterfDll = CreateBennerObject("BSINTERFACE0060.RotinasImportacaoSalario")
    vsRetornoMensagem = InterfDll.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    If vsRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If
  Else
    If Not InTransaction Then StartTransaction

    Set SQL = NewQuery
    SQL.Clear
    SQL.Add("UPDATE SAM_ROTINAIMPSAL SET SITUACAOIMPORTAR = '1' WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.ExecSQL

    If InTransaction Then Commit

    Set SQL = Nothing

    Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                 "BSBEN033", _
                                                 "CancelamentoSalario", _
                                                 "Cancelamento de salários", _
                                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                 "SAM_ROTINAIMPSAL", _
                                                 "SITUACAOIMPORTAR", _
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

	Set InterfDll = Nothing
  	Set dllBSServerExec = Nothing

  	WriteAudit("I", HandleOfTable("SAM_ROTINAIMPSAL"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
  		"Rotina de Atualização de Cadastro - Importação de salários")

	If VisibleMode Then
	  SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
	End If
  End If

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim InterfDll As Object
  Dim vsRetornoMensagem As Long
  Dim vsMensagemErro As String
  Dim viRetorno As Long
  Dim dllBSServerExec As Object

  If VisibleMode Then
    Set InterfDll = CreateBennerObject("BSINTERFACE0060.RotinasImportacaoSalario")
    vsRetornoMensagem = InterfDll.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    If vsRetornoMensagem = 1 Then
      bsShowMessage("Ocorreu erro no processo","I")
    End If
  Else
    Set dllBSServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")

    viRetorno = dllBSServerExec.ExecucaoImediata(CurrentSystem, _
                                                 "BSBEN033", _
                                                 "ProcessamentoSalario", _
                                                 "Processamento de salários", _
                                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                 "SAM_ROTINAIMPSAL", _
                                                 "SITUACAOPROCESSAR", _
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

	Set InterfDll = Nothing
  	Set dllBSServerExec = Nothing

  	WriteAudit("I", HandleOfTable("SAM_ROTINAIMPSAL"), CurrentQuery.FieldByName("HANDLE").AsInteger, _
  		"Rotina de Atualização de Cadastro - Importação de salários")

	If VisibleMode Then
	  SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
	End If
  End If
End Sub

Public Sub CONTRATOINICIAL_OnChange()
  CurrentQuery.FieldByName("CONTRATOFINAL").Clear
End Sub

Public Sub CONTRATOINICIAL_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  If (CurrentQuery.FieldByName("COMPETENCIA").IsNull) Then
    MsgBox("Primeiro informe a competência da rotina.")
  Else
    If (CurrentQuery.FieldByName("GRUPOCONTRATO").IsNull) Then
      MsgBox("Primeiro selecione um grupo de contratos.")
    Else
      Dim Interface     As Object
      Dim viHandle      As Long
      Dim vsCampos      As String
      Dim vsColunas     As String
      Dim vsCriterio    As String
      Dim vsNvl         As String
      Dim vsInfinito    As String
      Dim vsCompetencia As String

      vsColunas  = "CONTRATO|CONTRATANTE|DATACANCELAMENTO"
      vsCampos   = "Nº do Contrato|Contratante|Data de cancelamento"

      vsInfinito    = SQLAddYear(SQLDate(ServerDate),"200")
      vsCompetencia = SQLDate(CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
      If (InStr(SQLServer, "MSSQL") > 0) Then
        vsNvl         = "ISNULL"
      Else
        If (InStr(SQLServer,"ORACLE") > 0) Or (InStr(SQLServer,"CACHE") > 0) Then
          vsNvl         = "NVL"
        Else
          vsNvl         = "COALESCE"
        End If
      End If

      'Selecionar os contratos ativos (data de cancelamento nula ou maior/igual a competência da rotina) e que pertençam ao grupo de contratos selecionado.
      vsCriterio = ""
      vsCriterio = vsCriterio  + vsNvl + "(DATACANCELAMENTO, " + vsInfinito + ") >= " + vsCompetencia
      vsCriterio = vsCriterio  + " AND GRUPOCONTRATO = " + CurrentQuery.FieldByName("GRUPOCONTRATO").AsString

      Set Interface = CreateBennerObject("PROCURA.Procurar")
      viHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO", vsColunas, 1, vsCampos, vsCriterio, "Contratos", False, CONTRATOINICIAL.Text)

      If (viHandle > 0) Then
        CurrentQuery.Edit
        CurrentQuery.FieldByName("CONTRATOINICIAL").AsInteger = viHandle
      End If
      Set Interface = Nothing
    End If
  End If
End Sub

Public Sub CONTRATOFINAL_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  If (CurrentQuery.FieldByName("COMPETENCIA").IsNull) Then
    MsgBox("Primeiro informe a competência da rotina.")
  Else
    If (CurrentQuery.FieldByName("GRUPOCONTRATO").IsNull) Then
      MsgBox("Primeiro selecione um grupo de contratos.")
    Else
      If (CurrentQuery.FieldByName("CONTRATOINICIAL").IsNull) Then
        MsgBox("Primeiro defina o contrato inicial.")
      Else
        Dim Interface     As Object
        Dim viHandle      As Long
        Dim vsCampos      As String
        Dim vsColunas     As String
        Dim vsCriterio    As String
        Dim vsNvl         As String
        Dim vsInfinito    As String
        Dim vsCompetencia As String

        vsColunas  = "CONTRATO|CONTRATANTE|DATACANCELAMENTO"
        vsCampos   = "Nº do Contrato|Contratante|Data de cancelamento"


      vsInfinito    = SQLAddYear(SQLDate(ServerDate),"200")
      vsCompetencia = SQLDate(CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
      If (InStr(SQLServer, "MSSQL") > 0) Then
        vsNvl         = "ISNULL"
      Else
        If (InStr(SQLServer,"ORACLE") > 0) Or (InStr(SQLServer,"CACHE") > 0) Then
          vsNvl         = "NVL"
        Else
          vsNvl         = "COALESCE"
        End If
      End If


        Dim qContrato As Object
        Set qContrato = NewQuery

        qContrato.Clear
        qContrato.Add("SELECT CONTRATO        ")
        qContrato.Add("  FROM SAM_CONTRATO    ")
        qContrato.Add(" WHERE HANDLE = :HANDLE")
        qContrato.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATOINICIAL").AsString
        qContrato.Active = True

        'Selecionar os contratos ativos (data de cancelamento nula ou maior/igual a competência da rotina)
        'que possuam o código maior/igual ao contrato inicial e pertençam ao grupo de contratos selecionado.
        vsCriterio = ""
        vsCriterio = vsCriterio  + vsNvl + "(DATACANCELAMENTO, " + vsInfinito + ") >= " + vsCompetencia
        vsCriterio = vsCriterio  + " AND CONTRATO >= " + qContrato.FieldByName("CONTRATO").AsString
        vsCriterio = vsCriterio  + " AND GRUPOCONTRATO = " + CurrentQuery.FieldByName("GRUPOCONTRATO").AsString

        Set qContrato = Nothing

        Set Interface = CreateBennerObject("PROCURA.Procurar")
        viHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO", vsColunas, 1, vsCampos, vsCriterio, "Contratos", False, CONTRATOFINAL.Text)

        If (viHandle > 0) Then
          CurrentQuery.Edit
          CurrentQuery.FieldByName("CONTRATOFINAL").AsInteger = viHandle
        End If
        Set Interface = Nothing
      End If
    End If
  End If
End Sub

Public Sub GRUPOCONTRATO_OnChange()
  CurrentQuery.FieldByName("CONTRATOINICIAL").Clear
  CurrentQuery.FieldByName("CONTRATOFINAL").Clear
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vsMensagem As String

  If (CurrentQuery.FieldByName("TABFILTRO").AsInteger = 1) Then
    If ((Not CurrentQuery.FieldByName("CONTRATOINICIAL").IsNull) And (CurrentQuery.FieldByName("CONTRATOFINAL").IsNull)) Then
      vsMensagem = "Informe o contrato final do intervalo."
      If (VisibleMode) Then
        MsgBox(vsMensagem)
      Else
        CancelDescription = vsMensagem
      End If
      CanContinue = False
      Exit Sub
    End If

    If ((CurrentQuery.FieldByName("CONTRATOINICIAL").IsNull) And (Not CurrentQuery.FieldByName("CONTRATOFINAL").IsNull)) Then
      vsMensagem = "Informe o contrato inicial do intervalo."
      If (VisibleMode) Then
        MsgBox(vsMensagem)
      Else
        CancelDescription = vsMensagem
      End If
      CanContinue = False
      Exit Sub
    End If

    If ((Not CurrentQuery.FieldByName("CONTRATOINICIAL").IsNull) And (Not CurrentQuery.FieldByName("CONTRATOFINAL").IsNull)) Then
      If (CurrentQuery.FieldByName("CONTRATOINICIAL").AsInteger > CurrentQuery.FieldByName("CONTRATOFINAL").AsInteger) Then
        vsMensagem = "O código do contrato inicial não pode ser maior do que o código do contrato final."
        If (VisibleMode) Then
          MsgBox(vsMensagem)
        Else
          CancelDescription = vsMensagem
        End If
        CanContinue = False
        Exit Sub
      End If
    End If
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
    Case "BOTAOPROCESSAR"
	  BOTAOPROCESSAR_OnClick
    Case "BOTAOIMPORTAR"
      BOTAOIMPORTAR_OnClick
    Case "BOTAOCANCELAR"
      BOTAOCANCELAR_OnClick
  End Select
End Sub
