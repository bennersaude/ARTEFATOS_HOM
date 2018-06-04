'HASH: A24C262CF9B4CBFEADEC646CC6E06272
'Macro: SAM_MSPROCESSO
'#Uses "*UltimoDiaCompetencia"
'#Uses "*PrimeiroDiaCompetencia"
'#Uses "*bsShowMessage"

Public Sub BOTAOAGENDAMENTO_OnClick()
  Dim qr As Object
  Dim qr1 As Object
  Dim vSituacao As String
  Dim vTabela As String
  Dim vLegendaAgendamento As String
  Dim VLegendaAberta As String
  Dim vLegendaProcessada As String
  Set qr = NewQuery
  Set qr1 = NewQuery
  vTabela = "SAM_MSPROCESSO"
  vLegendaAgendamento = "D"
  VLegendaAberta = "A"
  vLegendaProcessada = "P"
  qr.Clear
  qr.Add("SELECT SITUACAO FROM " + vTabela + " WHERE HANDLE = :pHANDLE")
  qr.ParamByName("pHandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qr.Active = True
  vSituacao = qr.FieldByName("SITUACAO").AsString
  If vSituacao <> vLegendaAgendamento Then
    If ((vSituacao = VLegendaAberta) Or (vSituacao = vLegendaProcessada)) Then
      If bsShowMessage("Confirme o agendamento da rotina", "Q") = vbYes Then '(6=yes, 7=não)
        If Not InTransaction Then StartTransaction
        qr1.Clear
        qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
        qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qr1.ParamByName("pSituacao").AsString = vLegendaAgendamento
        qr1.ExecSQL
        If InTransaction Then Commit
      End If
    Else
      bsShowMessage("Rotina já foi processada e/ou gerada." + Chr(10) + _
      	  "Para alterar a situação, a rotina deverá ser CANCELADA antes!", "I")
    End If
  Else
    If bsShowMessage("Rotina já está agendada. Para retirar o agendamento pressione 'SIM'", "Q") = vbYes Then
      If Not InTransaction Then StartTransaction
      qr1.Clear
      qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
      qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      If CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
        qr1.ParamByName("pSituacao").AsString = VLegendaAberta
      Else
        qr1.ParamByName("pSituacao").AsString = vLegendaProcessada
      End If
      qr1.ExecSQL
      If InTransaction Then Commit
    End If
  End If
  Set qr = Nothing
  Set qr1 = Nothing
  If VisibleMode Then
    SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object
  Dim SQLOPERADORA As Object
  Set SQLOPERADORA = NewQuery
  Dim voperadora As Integer

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If bsShowMessage("Confirma o cancelamento?", "Q") = vbYes Then
    ' LOPES
    If Not CurrentQuery.FieldByName("USUARIOIMPORTACAO").IsNull Then
      SQLOPERADORA.Active = False
      SQLOPERADORA.Clear
      SQLOPERADORA.Add("SELECT COUNT(HANDLE) AS QT ")
      SQLOPERADORA.Add("  FROM SAM_MSPROCESSO_RETORNO_DET ")
      SQLOPERADORA.Add(" WHERE CABECALHO IN (SELECT HANDLE")
      SQLOPERADORA.Add("                       FROM SAM_MSPROCESSO_RETORNO_CAB")
      SQLOPERADORA.Add("                      WHERE CABECALHO = :HND)")
      SQLOPERADORA.Add("   AND SITUACAOREGISTRO = :SIT")
      SQLOPERADORA.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLOPERADORA.ParamByName("SIT").AsString = "2"
      SQLOPERADORA.Active = True
      If SQLOPERADORA.FieldByName("QT").AsInteger > 0 Then
        bsShowMessage("Existem registros que ainda não foram enviados. O cancelamento não é permitido.", "I")
        Exit Sub
      End If

	  If Not InTransaction Then StartTransaction
      SQLOPERADORA.Active = False
      SQLOPERADORA.Clear
      SQLOPERADORA.Add("DELETE ")
      SQLOPERADORA.Add("  FROM SAM_MSPROCESSO_RETORNO_DET ")
      SQLOPERADORA.Add(" WHERE CABECALHO IN (SELECT HANDLE")
      SQLOPERADORA.Add("                       FROM SAM_MSPROCESSO_RETORNO_CAB")
      SQLOPERADORA.Add("                      WHERE CABECALHO = :HND)")
      SQLOPERADORA.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLOPERADORA.ExecSQL
      If InTransaction Then Commit

	  If Not InTransaction Then StartTransaction
      SQLOPERADORA.Active = False
      SQLOPERADORA.Clear
      SQLOPERADORA.Add("DELETE ")
      SQLOPERADORA.Add("  FROM SAM_MSPROCESSO_RETORNO_CAB")
      SQLOPERADORA.Add(" WHERE CABECALHO = :HND")
      SQLOPERADORA.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLOPERADORA.ExecSQL
      If InTransaction Then Commit

	  If Not InTransaction Then StartTransaction
      SQLOPERADORA.Active = False
      SQLOPERADORA.Clear
      SQLOPERADORA.Add("UPDATE SAM_MSPROCESSO SET USUARIOIMPORTACAO = NULL, DATAIMPORTACAO = NULL")
      SQLOPERADORA.Add(" WHERE HANDLE = :HND")
      SQLOPERADORA.ParamByName("HND").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLOPERADORA.ExecSQL
      If InTransaction Then Commit
    Else
      If CurrentQuery.FieldByName("tabtipooperadora").AsInteger = 1 Then
        voperadora = CurrentQuery.FieldByName("OPERADORA").AsInteger

        SQLOPERADORA.Active = False
        SQLOPERADORA.Clear
        SQLOPERADORA.Add("SELECT MAX(DATAFINAL)DATAFINAL     ")
        SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
        SQLOPERADORA.Add(" WHERE OPERADORA = :OPERADORA      ")
        SQLOPERADORA.Add("   AND TABTIPOOPERADORA = :TABTIPOOPERADORA ")
        SQLOPERADORA.Add("   AND SITUACAO  = 'P'             ")
        SQLOPERADORA.Add("   AND HANDLE  <> :HANDLE          ")
        SQLOPERADORA.Add("   AND DATAIMPORTACAO IS NULL      ")
        SQLOPERADORA.Add("   AND TABTIPO = 1                 ")
        SQLOPERADORA.ParamByName("OPERADORA").Value = voperadora
        SQLOPERADORA.ParamByName("TABTIPOOPERADORA").Value = CurrentQuery.FieldByName("tabtipooperadora").AsInteger
        SQLOPERADORA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQLOPERADORA.Active = True

        If SQLOPERADORA.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
          bsShowMessage("Existem Processos fechados com mesma operadora com data final superior. Não é possível cancelar este Processo.", "E")
          CanContinue = False
          Exit Sub
        End If

        SQLOPERADORA.Active = False
        SQLOPERADORA.Clear
        SQLOPERADORA.Add("SELECT HANDLE                      ")
        SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
        SQLOPERADORA.Add(" WHERE OPERADORA = :OPERADORA      ")
        SQLOPERADORA.Add("   AND TABTIPOOPERADORA = :TABTIPOOPERADORA ")
        SQLOPERADORA.Add("   AND SITUACAO  = 'A'             ")
        SQLOPERADORA.Add("   AND DATAIMPORTACAO IS NULL      ")
        SQLOPERADORA.Add("   AND TABTIPO = 1                 ")
        SQLOPERADORA.ParamByName("OPERADORA").Value = voperadora
        SQLOPERADORA.ParamByName("TABTIPOOPERADORA").Value = CurrentQuery.FieldByName("tabtipooperadora").AsInteger
        SQLOPERADORA.Active = True

        If Not SQLOPERADORA.EOF Then
          bsShowMessage("Existem Processos abertos com esta operadora. Não é possível cancelar este Processo.", "E")
          CanContinue = False
          Exit Sub
        End If
      Else
        voperadora = CurrentQuery.FieldByName("OPERADORAADM").AsInteger

        SQLOPERADORA.Active = False
        SQLOPERADORA.Clear
        SQLOPERADORA.Add("SELECT MAX(DATAFINAL)DATAFINAL     ")
        SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
        SQLOPERADORA.Add(" WHERE OPERADORAADM = :OPERADORA   ")
        SQLOPERADORA.Add("   AND TABTIPOOPERADORA = :TABTIPOOPERADORA ")
        SQLOPERADORA.Add("   AND SITUACAO  = 'P'             ")
        SQLOPERADORA.Add("   AND HANDLE  <> :HANDLE          ")
        SQLOPERADORA.Add("   AND DATAIMPORTACAO IS NULL      ")
        SQLOPERADORA.Add("   AND TABTIPO = 1                 ")
        SQLOPERADORA.ParamByName("OPERADORA").Value = voperadora
        SQLOPERADORA.ParamByName("TABTIPOOPERADORA").Value = CurrentQuery.FieldByName("tabtipooperadora").AsInteger
        SQLOPERADORA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQLOPERADORA.Active = True

        If SQLOPERADORA.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
          bsShowMessage("Existem Processos fechados com mesma operadora com data final superior. Não é possível cancelar este Processo.", "E")
          CanContinue = False
          Exit Sub
        End If
        SQLOPERADORA.Active = False
        SQLOPERADORA.Clear
        SQLOPERADORA.Add("SELECT HANDLE                      ")
        SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
        SQLOPERADORA.Add(" WHERE OPERADORAADM = :OPERADORA   ")
        SQLOPERADORA.Add("   AND TABTIPOOPERADORA = :TABTIPOOPERADORA ")
        SQLOPERADORA.Add("   AND SITUACAO  = 'A'             ")
        SQLOPERADORA.Add("   AND DATAIMPORTACAO IS NULL      ")
        SQLOPERADORA.ParamByName("OPERADORA").Value = voperadora
        SQLOPERADORA.ParamByName("TABTIPOOPERADORA").Value = CurrentQuery.FieldByName("tabtipooperadora").AsInteger
        SQLOPERADORA.Active = True

        If Not SQLOPERADORA.EOF Then
          bsShowMessage("Existem Processos abertos com esta operadora. Não é possível cancelar este Processo.", "E")
          CanContinue = False
          Exit Sub
        End If
      End If
    End If

    'SQL.Active=False
    'Set SQL =Nothing
    'SqlOperadora =False
    Set SQLOPERADORA = Nothing

    Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
    Obj.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Obj = Nothing

    If Not WebMode Then
	  RefreshNodesWithTable("SAM_MSPROCESSO")
	End If

    WriteAudit("C", HandleOfTable("SAM_MSPROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de registro no Ministério da Saúde - Cancelamento da Rotina")
  End If
End Sub

Public Sub BOTAODEVOLUCAO_OnClick()
  ' lopes
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("USUARIOIMPORTACAO").IsNull Then
    bsShowMessage("A rotina já foi importada", "I")
    Exit Sub
  End If

  Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
  Obj.Importar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
  If Not WebMode Then
    RefreshNodesWithTable("SAM_MSPROCESSO")
  End If

  WriteAudit("P", HandleOfTable("SAM_MSPROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de registro no Ministério da Saúde - Importação")
End Sub

Public Sub BOTAOEXCLUIR_OnClick()
  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If bsShowMessage("Confirma a exclusão do Processo ?", "Q") = vbYes Then
    If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
      Dim SQL As Object
      Set SQL = NewQuery

      SQL.Add("SELECT COUNT(HANDLE) QTDPROCESSOS")
      SQL.Add("FROM SAM_MSPROCESSO")
      SQL.Add("WHERE HANDLE > :HMSPROCESSO")
      SQL.Add("  AND SITUACAO = 'P'")
      SQL.ParamByName("HMSPROCESSO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.Active = True

      If SQL.FieldByName("QTDPROCESSOS").AsInteger <2 Then
        bsShowMessage("O último e o penúltimo Processo com status de 'Processado' não podem ser excluídos", "I")
        Exit Sub
      End If

      SQL.Active = False
      Set SQL = Nothing
    End If

    Dim Obj As Object

    Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
    Obj.Excluir(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Obj = Nothing

    If Not WebMode Then
      RefreshNodesWithTable("SAM_MSPROCESSO")
    End If

    WriteAudit("E", HandleOfTable("SAM_MSPROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de registro no Ministério da Saúde - Exclusão da Rotina")
  End If
End Sub

Public Sub BOTAOEXPORTAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.FieldByName("SITUACAO").AsString = "A" Then
    bsShowMessage("A rotina não foi processada.", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
    bsShowMessage("A rotina não foi processada", "I")
    Exit Sub
  End If

  Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
  Obj.Exportar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
  If Not WebMode Then
    RefreshNodesWithTable("SAM_MSPROCESSO")
  End If

  WriteAudit("E", HandleOfTable("SAM_MSPROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de registro no Ministério da Saúde - Exportação do arquivo")

End Sub

Public Sub BOTAOIMPORTARCONFERENCIA_OnClick()
' Roberto - SMS 77918
  Dim Obj As Object

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(HANDLE) QT ")
  SQL.Add("  FROM SAM_MSPROCESSO   ")
  SQL.Add(" WHERE HANDLE <> :HANDLE ")
  SQL.Add("   AND SITUACAO  = 'P' ")
  SQL.Add("   AND TABTIPO = 3  ")
  SQL.Add("   AND DATAFINAL BETWEEN :DATAINI AND :DATAFIM ")

  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("DATAINI").AsDateTime = PrimeiroDiaCompetencia(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)
  SQL.ParamByName("DATAFIM").AsDateTime = UltimoDiaCompetencia(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

  SQL.Active = True
  If SQL.FieldByName("QT").AsInteger > 0 Then
	If Not WebMode Then
	  DATAFINAL.SetFocus
	End If
  	bsShowMessage("Uma rotina SIB de Conferência com esta competência já foi processada! Verifique a data final informada.", "I")

    SQL.Active = False
    Set SQL = Nothing
	Exit Sub
  End If

  SQL.Active = False
  Set SQL = Nothing

  Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
  Obj.ImportarArquivoConferencia(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing

  If Not WebMode Then
    RefreshNodesWithTable("SAM_MSPROCESSO")
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("USUARIOPROCESSAMENTO").IsNull Then
    bsShowMessage("A rotina já foi processada", "I")
    Exit Sub
  End If

  Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
  Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
  If Not WebMode Then
    RefreshNodesWithTable("SAM_MSPROCESSO")
  End If

  WriteAudit("P", HandleOfTable("SAM_MSPROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de registro no Ministério da Saúde - Processamento da Rotina")
End Sub

Public Sub BOTAOVERIFICAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Set Obj = CreateBennerObject("SAMMSPROCESSO.Geral")
  Obj.Verificar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing

  WriteAudit("V", HandleOfTable("SAM_MSPROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de registro no Ministério da Saúde - Verificação do Cadastro Atual")

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    DATAFINAL.ReadOnly = True
  Else
    DATAFINAL.ReadOnly = False
  End If

  'Roberto SMS 77918
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 Then  'Conferência SIB
    BOTAOIMPORTARCONFERENCIA.Enabled = True
    BOTAOCANCELAR.Enabled = False
    BOTAOVERIFICAR.Enabled = False
    BOTAOPROCESSAR.Enabled = False
    BOTAODEVOLUCAO.Enabled = False
    BOTAOEXPORTAR.Enabled = False
    If (CurrentQuery.FieldByName("SITUACAO").AsString <> "P") And (CurrentQuery.FieldByName("SITUACAO").AsString <> "S") Then
      BOTAOCANCELAR.Enabled = False
      BOTAOEXCLUIR.Enabled  = True
    Else
      BOTAOCANCELAR.Enabled = True
      BOTAOEXCLUIR.Enabled  = False
    End If
  ElseIf CurrentQuery.FieldByName("TABTIPO").AsInteger = 2 Then  'Retorno SIB
    BOTAOPROCESSAR.Enabled = False
    BOTAOIMPORTARCONFERENCIA.Enabled = False
    If (CurrentQuery.FieldByName("SITUACAO").AsString <> "P") Then
      BOTAOCANCELAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "S")
      BOTAOEXPORTAR.Enabled = False
      BOTAOEXCLUIR.Enabled  = (CurrentQuery.FieldByName("SITUACAO").AsString <> "S")
      BOTAODEVOLUCAO.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString <> "S")
    Else
      BOTAOCANCELAR.Enabled = True
      BOTAOEXPORTAR.Enabled = True
      BOTAOEXCLUIR.Enabled  = False
      BOTAODEVOLUCAO.Enabled = False
    End If
  Else ' Remessa SIB
    BOTAODEVOLUCAO.Enabled = False
    BOTAOIMPORTARCONFERENCIA.Enabled = False
    If (CurrentQuery.FieldByName("SITUACAO").AsString <> "P") Then
      BOTAOPROCESSAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString <> "S")
      BOTAOCANCELAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "S")
      BOTAOEXCLUIR.Enabled  = (CurrentQuery.FieldByName("SITUACAO").AsString <> "S")
      BOTAOEXPORTAR.Enabled  = False
    Else
      BOTAOPROCESSAR.Enabled = False
      BOTAOCANCELAR.Enabled = True
      BOTAOEXCLUIR.Enabled  = False
      BOTAOEXPORTAR.Enabled  = True
    End If
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLOPERADORA As Object
  Set SQLOPERADORA = NewQuery
  Dim viTipoOperadora As Integer
  Dim viOperadora As Integer

  ' SMS 77918
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then
    ' Remessa
    If CurrentQuery.FieldByName("DIRETORIODESTINO").IsNull Then
      If Not WebMode Then
        ARQUIVOORIGEM.SetFocus
      End If
      bsShowMessage("É necessário informar Diretório de Destino na rotina de Remessa do SIB!", "E")
      CanContinue = False
      Set SQLOPERADORA = Nothing
      Exit Sub
    End If

    viTipoOperadora = CurrentQuery.FieldByName("TABTIPOOPERADORA").AsInteger

    If viTipoOperadora = 1 Then
      If CurrentQuery.FieldByName("OPERADORA").IsNull Then
        If Not WebMode Then
          OPERADORA.SetFocus
        End If
        bsShowMessage("O campo 'Operadora' deve ser preenchido!", "E")
        CanContinue = False
        Set SQLOPERADORA = Nothing
        Exit Sub
      End If

      viOperadora = CurrentQuery.FieldByName("OPERADORA").AsInteger
    Else
      If CurrentQuery.FieldByName("OPERADORAADM").IsNull Then
        If Not WebMode Then
          OPERADORAADM.SetFocus
        End If
        bsShowMessage("O campo 'Operadora' deve ser preenchido!", "E")
        CanContinue = False
        Set SQLOPERADORA = Nothing
        Exit Sub
      End If

      viOperadora = CurrentQuery.FieldByName("OPERADORAADM").AsInteger
    End If

    SQLOPERADORA.Active = False
    SQLOPERADORA.Clear
    SQLOPERADORA.Add("SELECT HANDLE                      ")
    SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
    SQLOPERADORA.Add(" WHERE TABTIPO = 1                 ")
    SQLOPERADORA.Add("   AND DATAFINAL BETWEEN :DATAINICIAL AND :DATAFINAL ")
    SQLOPERADORA.Add("   AND HANDLE  <> :HANDLE                ")
    SQLOPERADORA.Add("   AND TABTIPOOPERADORA = :TIPOOPERADORA ")

    If viTipoOperadora = 1 Then
      SQLOPERADORA.Add("   AND OPERADORA = :OPERADORA      ")
    Else
      SQLOPERADORA.Add("   AND OPERADORAADM = :OPERADORA   ")
    End If

    SQLOPERADORA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQLOPERADORA.ParamByName("TIPOOPERADORA").AsInteger = viTipoOperadora
    SQLOPERADORA.ParamByName("OPERADORA").AsInteger     = viOperadora
    SQLOPERADORA.ParamByName("DATAINICIAL").AsDateTime  = vdDataInicial
    SQLOPERADORA.ParamByName("DATAFINAL").AsDateTime    = vdDataFinal
    SQLOPERADORA.Active =True

    If Not SQLOPERADORA.FieldByName("HANDLE").IsNull Then
      If Not WebMode Then
        DATAFINAL.SetFocus
      End If
      bsShowMessage("Já existe Rotina de Remessa SIB nesta competência! Verifique a data final.", "E")
      CanContinue = False
      Set SQLOPERADORA = Nothing
      Exit Sub
    End If
  ElseIf CurrentQuery.FieldByName("TABTIPO").AsInteger = 2 Then
    ' Retorno

    If CurrentQuery.FieldByName("ARQUIVOORIGEM").IsNull Then
      If Not WebMode Then
        ARQUIVOORIGEM.SetFocus
      End If
      bsShowMessage("É necessário informar Arquivo Devolução na rotina de Retorno do SIB!", "E")
      CanContinue = False
      Set SQLOPERADORA = Nothing
      Exit Sub
    End If

    SQLOPERADORA.Active =False
    SQLOPERADORA.Clear
    SQLOPERADORA.Add("SELECT HANDLE                      ")
    SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
    SQLOPERADORA.Add(" WHERE TABTIPO = 2                 ")
    SQLOPERADORA.Add("   AND DATAFINAL BETWEEN :DATAINICIAL AND :DATAFINAL ")
    SQLOPERADORA.Add("   AND HANDLE  <> :HANDLE                ")

    SQLOPERADORA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQLOPERADORA.ParamByName("DATAINICIAL").AsDateTime  = vdDataInicial
    SQLOPERADORA.ParamByName("DATAFINAL").AsDateTime    = vdDataFinal
    SQLOPERADORA.Active = True

    If Not SQLOPERADORA.FieldByName("HANDLE").IsNull Then
      If Not WebMode Then
        DATAFINAL.SetFocus
      End If
      bsShowMessage("Já existe Rotina de Retorno SIB nesta competência! Verifique a data final.", "E")
      CanContinue = False
      Set SQLOPERADORA = Nothing
      Exit Sub
    End If
  ElseIf CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 Then
    ' Conferência

    If CurrentQuery.FieldByName("DIRETORIODESTINOCONFERENCIA").IsNull Then
      If Not WebMode Then
        ARQUIVOORIGEM.SetFocus
      End If
      bsShowMessage("É necessário informar o Local Arquivo Conferência na rotina de Conferência do SIB!", "E")
      CanContinue = False
      Set SQLOPERADORA = Nothing
      Exit Sub
    End If

    SQLOPERADORA.Active = False
    SQLOPERADORA.Clear
    SQLOPERADORA.Add("SELECT HANDLE                      ")
    SQLOPERADORA.Add("  FROM SAM_MSPROCESSO              ")
    SQLOPERADORA.Add(" WHERE TABTIPO = 3                 ")
    SQLOPERADORA.Add("   AND DATAFINAL BETWEEN :DATAINICIAL AND :DATAFINAL ")
    SQLOPERADORA.Add("   AND HANDLE <> :HANDLE           ")

    SQLOPERADORA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQLOPERADORA.ParamByName("DATAINICIAL").AsDateTime  = vdDataInicial
    SQLOPERADORA.ParamByName("DATAFINAL").AsDateTime    = vdDataFinal
    SQLOPERADORA.Active = True

    If Not SQLOPERADORA.FieldByName("HANDLE").IsNull Then
      If Not WebMode Then
        DATAFINAL.SetFocus
      End If
      bsShowMessage("Já existe Rotina de Conferência SIB nesta competência! Verifique a data final.", "E")
      CanContinue = False
      Set SQLOPERADORA = Nothing
      Exit Sub
    End If
  Else
    bsShowMessage("Deve ser informado um tipo para a rotina SIB!", "E")
    CanContinue = False
    Set SQLOPERADORA = Nothing
    Exit Sub
  End If

  Set SQLOPERADORA = Nothing

  ' fim SMS 77918
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BOTAOAGENDAMENTO" Then
	BOTAOAGENDAMENTO_OnClick
  ElseIf CommandID = "BOTAOCANCELAR" Then
	BOTAOCANCELAR_OnClick
  ElseIf CommandID = "BOTAODEVOLUCAO" Then
	BOTAODEVOLUCAO_OnClick
  ElseIf CommandID = "BOTAOEXCLUIR" Then
	BOTAOEXCLUIR_OnClick
  ElseIf CommandID = "BOTAOEXPORTAR" Then
	BOTAOEXPORTAR_OnClick
  ElseIf CommandID = "BOTAOIMPORTARCONFERENCIA" Then
	BOTAOIMPORTARCONFERENCIA_OnClick
  ElseIf CommandID = "BOTAOPROCESSAR" Then
	BOTAOPROCESSAR_OnClick
  ElseIf CommandID = "BOTAOVERIFICAR" Then
	BOTAOVERIFICAR_OnClick
  End If
End Sub
