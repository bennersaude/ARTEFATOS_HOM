'HASH: 6CA1E3019C9B4147F0BD2DBDE6169A6A
'TABELA AEX_EMPCONECT
' atualizada em 10/08/2007
'#Uses "*bsShowMessage"

Public Sub CamposTabelasNaoParam
  CurrentQuery.FieldByName("BOOLBENEFICIARIO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIABENEF").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIACONTRAT").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIAEVENT").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIAS").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCOBERTCONTRAT").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCOBERTURABENEF").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCONTRATO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITBENEF").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITCONTRAT").AsBoolean = False
  CurrentQuery.FieldByName("BOOLMODBENEF").AsBoolean = False
  CurrentQuery.FieldByName("BOOLMODCONTRATOS").AsBoolean = False
  CurrentQuery.FieldByName("BOOLESPPRESTA").AsBoolean = False
  CurrentQuery.FieldByName("BOOLEVENTTGECBHPM").AsBoolean = False
  CurrentQuery.FieldByName("BOOLGRUPOESP").AsBoolean = False
  CurrentQuery.FieldByName("BOOLGRUPOESPEVENTO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMMUNICIPIO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMPESTADO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMPGERAL").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMPPRESTA").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITEEVENTO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITES").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLOCAPRESTA").AsBoolean = False
  CurrentQuery.FieldByName("BOOLPRESTATODOCAD").AsBoolean = False
  CurrentQuery.FieldByName("BOOLREGRAEXCESSAO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCONTLIMITBENEF").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCONTRATOMOD").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond
  CurrentQuery.FieldByName("BOOLNIVELAUTORIZACAO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLNIVELEVENTO").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond - Fim
End Sub

Public Sub CamposTabelasNaoParamArq
  CurrentQuery.FieldByName("BOOLBENEFICIARIOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIABENEFARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIACONTRATARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIAEVENTARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCARENCIASARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCOBERTCONTRATARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCOBERTURABENEFARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCONTRATOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITBENEFARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITCONTRATARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLMODBENEFARQ").AsBoolean = False
  '	CurrentQuery.FieldByName("BOOLMODCONTRATOSARQ").AsBoolean    = False
  CurrentQuery.FieldByName("BOOLESPPRESTAARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLEVENTTGECBHPMARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLGRUPOESPARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLGRUPOESPEVENTOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMMUNICIPIOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMPESTADOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMPGERALARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLINCOMPPRESTAARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITEEVENTOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLIMITESARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLLOCAPRESTAARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLPRESTATODOCADARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLREGRAEXCESSAOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCONTLIMITBENEFARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLCONTRATOMODARQ").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond
  CurrentQuery.FieldByName("BOOLNIVELAUTORIZACAOARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOLLNIVELEVENTOARQ").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond = Fim
End Sub


Public Sub CamposTabelasParam
  CurrentQuery.FieldByName("BOOLNEGACAO").AsBoolean = False
  CurrentQuery.FieldByName("BOOLESPECIALIDADES").AsBoolean = False
  CurrentQuery.FieldByName("BOOLPRESTADORCONECT").AsBoolean = False
  CurrentQuery.FieldByName("BOOLTIPOPRESTA").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond
  CurrentQuery.FieldByName("BOOLNIVELCONTRATOEVENTO").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond - Fim
End Sub

Public Sub CamposTabelasParamArq
  CurrentQuery.FieldByName("BOOLESPECIALIDADESARQ").AsBoolean = False
  CurrentQuery.FieldByName("BOOLPRESTADORCONECTARQ").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond
  CurrentQuery.FieldByName("BOOLNIVELCONTRATOEVENTOARQ").AsBoolean = False
  ' --- SMS - 75160 --- 18/01/2007 - Drummond - Fim
End Sub

Public Function VerificaLogs() As Boolean
  Dim qVerificaTab1 As Object
  Dim qVerificaTab2 As Object

  VerificaLogs = False
  Set qVerificaTab1 = NewQuery
  qVerificaTab1.Clear
  qVerificaTab1.Add("SELECT COUNT(1) QTD              ")
  qVerificaTab1.Add("  FROM AEX_LOGCARGATABELAS       ")
  qVerificaTab1.Add(" WHERE STATUSPROCESSO = 'N'      ")
  qVerificaTab1.Add("   AND EMPRESAEMS = :EMPRESAEMS  ")
  qVerificaTab1.ParamByName("EMPRESAEMS").Value = CurrentQuery.FieldByName("HANDLE").Value
  qVerificaTab1.Active = True
  If qVerificaTab1.FieldByName("QTD").AsInteger > 0 Then
    VerificaLogs = True
  End If
  Set qVerificaTab1 = Nothing

  Set qVerificaTab2 = NewQuery
  qVerificaTab2.Clear
  qVerificaTab2.Add("SELECT COUNT(1) QTD")
  qVerificaTab2.Add("  FROM AEX_LOGENVIO")
  qVerificaTab2.Add(" WHERE PROCESSADO = 'N'")
  qVerificaTab2.Add("   AND EMPCONECT = :EMPCONECT")
  qVerificaTab2.ParamByName("EMPCONECT").Value = CurrentQuery.FieldByName("HANDLE").Value
  qVerificaTab2.Active = True
  If qVerificaTab2.FieldByName("QTD").AsInteger > 0 Then
    VerificaLogs = True
  End If
  Set qVerificaTab2 = Nothing
End Function



Public Function VerificaTabela(psTabela As String) As Boolean
  ' ******** INICIO SMS - 39471 - 06/07/2005 - DRUMMOND ********
  Dim QVeriTabela As Object

  VerificaTabela = True

  Set QVeriTabela = NewQuery
  QVeriTabela.Clear
  QVeriTabela.Add("SELECT COUNT(1) QTD")
  QVeriTabela.Add("  FROM AEX_INFOTABELAS")
  QVeriTabela.Add(" WHERE NOME = :NOME")
  QVeriTabela.Add("   AND EMPCONECT = :EMPCO")
  QVeriTabela.ParamByName("NOME").AsString = psTabela
  QVeriTabela.ParamByName("EMPCO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QVeriTabela.Active = True

  If QVeriTabela.FieldByName("QTD").AsInteger = 0 Then
    VerificaTabela = False
  End If
  Set QVeriTabela = Nothing
  ' ******** FIM    SMS - 39471 - 06/07/2005 - DRUMMOND ********
End Function


Public Sub BOOLREPROCESSARARQ_OnClick()
  Dim obj As Object
  Dim SPP As Object
  '--------------------------------------------------------
  'Alterando o campo processado das tabelas AEX
  '--------------------------------------------------------
  Set SPP = NewStoredProc
  SPP.AutoMode = True
  SPP.Name = "BSAEX_REPROCESSARARQUIVO"
  SPP.ExecProc
  SPP.AutoMode = False
  '--------------------------------------------------------
  'Fim da alteração
  '--------------------------------------------------------
  'Chamando o processo de gerar arquivo
  '--------------------------------------------------------
  Set obj = CreateBennerObject("BSAte005.Rotinas")
  obj.GerarArquivo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set obj = Nothing

End Sub

Public Sub BOTAOBAIXARARQUIVOS_OnClick()
	Dim BSAte008 As Object

	Set BSAte008 = CreateBennerObject("BSAte008.Rotinas")
	BSAte008.BaixaArquivos(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser)
	Set BSAte008 = Nothing
End Sub

Public Sub BOTAOCAMINHO_OnClick()
  Dim Interface As Object
  Dim vPath As String


  Set Interface = CreateBennerObject("BSPRE001.Rotinas")
  vPath = Interface.SelecionarDiretorio(CurrentSystem)

  If vPath <>"" Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CAMINHOPARAARQUIVO").AsString = vPath
  End If

End Sub

Public Sub BOTAOCORRIGELOGCARGATABELAS_OnClick()
  Dim qUpdateCarga As Object
  Dim qNomeUsu As Object

  If bsshowmessage("Atenção!Ao mudar o status de um ou mais processos poderá causar problemas em processo em andamento se existir. Tem certeza que deseja continuar?", "Q") Then
    InfoDescription = "Teste"
    Exit Sub
    Set qNomeUsu = NewQuery
    qNomeUsu.Clear
    qNomeUsu.Add("SELECT NOME FROM Z_GRUPOUSUARIOS ")
    qNomeUsu.Add(" WHERE HANDLE = :HANDLE          ")
    qNomeUsu.ParamByName("HANDLE").Value = CurrentUser
    qNomeUsu.Active = True

	If Not InTransaction Then StartTransaction

    Set qUpdateCarga = NewQuery
    qUpdateCarga.Clear
    qUpdateCarga.Add("UPDATE AEX_LOGCARGATABELAS SET                       ")
    qUpdateCarga.Add("       STATUSPROCESSO = 'S',                         ")
    qUpdateCarga.Add("       OBSERVACOES = :OBSERVACOES,                   ")
    qUpdateCarga.Add("       FIMEXECUCAO = :FIMEXECUCAO                    ")
    'qUpdateCarga.Add("       TEMPOEXECUCAO = :FIMEXECUCAO1 - INICIOEXECUCAO ")
    qUpdateCarga.Add(" WHERE STATUSPROCESSO = 'N'                          ")
    qUpdateCarga.ParamByName("FIMEXECUCAO").Value = ServerNow
    'qUpdateCarga.ParamByName("FIMEXECUCAO1").AsDateTime = ServerNow
    If qNomeUsu.FieldByName("NOME").AsString = "" Then
      qUpdateCarga.ParamByName("OBSERVACOES").Value = "Status do log modificado através do botão de correção do log. Usuário: " + Str(CurrentUser)
    Else
      qUpdateCarga.ParamByName("OBSERVACOES").Value = "Status do log modificado através do botão de correção do log. Usuário: " + Str(CurrentUser) + " - " + qNomeUsu.FieldByName("NOME").AsString
    End If
    qUpdateCarga.ExecSQL

    If InTransaction Then Commit

    Set qNomeUsu = Nothing
    Set qUpdateCarga = Nothing
  End If
End Sub

Public Sub BOTAOCORRIGELOGERAARQUIVO_OnClick()
  Dim qUpdateArquivo As Object
  Dim qNomeUsu As Object

  If MsgBox("Atenção!Ao mudar o status de um ou mais processos poderá causar problemas em processo em andamento se existir. Tem certeza que deseja continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    Set qNomeUsu = NewQuery
    qNomeUsu.Clear
    qNomeUsu.Add("SELECT NOME FROM Z_GRUPOUSUARIOS ")
    qNomeUsu.Add(" WHERE HANDLE = :HANDLE          ")
    qNomeUsu.ParamByName("HANDLE").Value = CurrentUser
    qNomeUsu.Active = True

    If Not InTransaction Then StartTransaction

    Set qUpdateArquivo = NewQuery
    qUpdateArquivo.Clear
    qUpdateArquivo.Add("UPDATE AEX_LOGENVIO SET           ")
    qUpdateArquivo.Add("       PROCESSADO = 'S',          ")
    qUpdateArquivo.Add("       OBSERVACAO = :OBSERVACAO,  ")
    qUpdateArquivo.Add("       FIMEXECUCAO = :FIMEXECUCAO ")
    qUpdateArquivo.Add(" WHERE PROCESSADO = 'N'           ")
    qUpdateArquivo.ParamByName("FIMEXECUCAO").Value = ServerNow
    If qNomeUsu.FieldByName("NOME").AsString = "" Then
      qUpdateArquivo.ParamByName("OBSERVACAO").Value = "Status do log modificado através do botão de correção do log. Usuário: " + Str(CurrentUser)
    Else
      qUpdateArquivo.ParamByName("OBSERVACAO").Value = "Status do log modificado através do botão de correção do log. Usuário: " + Str(CurrentUser) + " - " + qNomeUsu.FieldByName("NOME").AsString
    End If
    qUpdateArquivo.ExecSQL

    If InTransaction Then Commit

    Set qNomeUsu = Nothing
    Set qUpdateArquivo = Nothing
  End If
End Sub

Public Sub BOTAOGERAARQ_OnClick()
  Dim obj As Object
  Dim vbTabela As Boolean
  'Verifico se o registro está em edição
  If CurrentQuery.State = 1 Then

    vbTabela = True

    ' --- SMS - 75160 --- 18/01/2007 - Drummond
    If (CurrentQuery.FieldByName("BOOLNIVELAUTORIZACAOARQ").AsBoolean) And (Not VerificaTabela("AEX_NIVELAUTORIZACAO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLNIVELAUTORIZACAOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLNIVELEVENTOARQ").AsBoolean) And (Not VerificaTabela("AEX_NIVELAUTORIZACAOEVENTO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLNIVELEVENTOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLNIVELCONTRATOEVENTOARQ").AsBoolean) And (Not VerificaTabela("AEX_NIVELCONTRATOEVENTO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLNIVELCONTRATOEVENTOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If
    ' --- SMS - 75160 --- 18/01/2007 - Drummond - Fim

    If (CurrentQuery.FieldByName("BOOLBENEFICIARIOARQ").AsBoolean) And (Not VerificaTabela("AEX_BENEFICIARIO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLBENEFICIARIOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCARENCIABENEFARQ").AsBoolean) And (Not VerificaTabela("AEX_BENEFICIARIO_CARENCIA")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCARENCIABENEFARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCARENCIACONTRATARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTRATO_CARENCIA")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCARENCIACONTRATARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCARENCIAEVENTARQ").AsBoolean) And (Not VerificaTabela("AEX_CARENCIA_EVENTO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCARENCIAEVENTARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCARENCIASARQ").AsBoolean) And (Not VerificaTabela("AEX_CARENCIA")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCARENCIASARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCOBERTCONTRATARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTRATO_COBERTURA")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCOBERTCONTRATARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCOBERTURABENEFARQ").AsBoolean) And (Not VerificaTabela("AEX_BENEFICIARIO_COBERTURA")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCOBERTURABENEFARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCONTRATOARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTRATO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCONTRATOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLLIMITBENEFARQ").AsBoolean) And (Not VerificaTabela("AEX_BENEFICIARIO_LIMITE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLLIMITBENEFARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLLIMITCONTRATARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTRATO_LIMITE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLLIMITCONTRATARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLMODBENEFARQ").AsBoolean) And (Not VerificaTabela("AEX_BENEFICIARIO_PLANO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLMODBENEFARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    '		If (CurrentQuery.FieldByName("BOOLMODCONTRATOSARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTRATO_MOD")) Then
    '			CurrentQuery.Edit
    '			CurrentQuery.FieldByName("BOOLMODCONTRATOSARQ").AsBoolean = False
    '			CurrentQuery.Post
    '			vbTabela = False
    '		End If

    If (CurrentQuery.FieldByName("BOOLESPPRESTAARQ").AsBoolean) And (Not VerificaTabela("AEX_PRESTADOR_SUBESPECIALIDADE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLESPPRESTAARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLEVENTTGECBHPMARQ").AsBoolean) And (Not VerificaTabela("AEX_TGE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLEVENTTGECBHPMARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLGRUPOESPARQ").AsBoolean) And (Not VerificaTabela("AEX_SUBESPECIALIDADE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLGRUPOESPARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLGRUPOESPEVENTOARQ").AsBoolean) And (Not VerificaTabela("AEX_PADRAO_SUBESPECIALIDADE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLGRUPOESPEVENTOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLINCOMMUNICIPIOARQ").AsBoolean) And (Not VerificaTabela("AEX_INCOMP_MUNICIPIO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLINCOMMUNICIPIOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLINCOMPESTADOARQ").AsBoolean) And (Not VerificaTabela("AEX_INCOMP_ESTADO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLINCOMPESTADOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLINCOMPGERALARQ").AsBoolean) And (Not VerificaTabela("AEX_INCOMP_GERAL")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLINCOMPGERALARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLINCOMPPRESTAARQ").AsBoolean) And (Not VerificaTabela("AEX_INCOMP_PRESTADOR")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLINCOMPPRESTAARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLLIMITEEVENTOARQ").AsBoolean) And (Not VerificaTabela("AEX_LIMITE_EVENTO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLLIMITEEVENTOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLLIMITESARQ").AsBoolean) And (Not VerificaTabela("AEX_LIMITE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLLIMITESARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLLOCAPRESTAARQ").AsBoolean) And (Not VerificaTabela("AEX_PRESTADOR_LOCALIDADE")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLLOCAPRESTAARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLPRESTATODOCADARQ").AsBoolean) And (Not VerificaTabela("AEX_PRESTADOR_CRD")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLPRESTATODOCADARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLREGRAEXCESSAOARQ").AsBoolean) And (Not VerificaTabela("AEX_PRESTADOR_REGRAEXCESSAO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLREGRAEXCESSAOARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCONTLIMITBENEFARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTAGEM_BENEFICIARIO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCONTLIMITBENEFARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    If (CurrentQuery.FieldByName("BOOLCONTRATOMODARQ").AsBoolean) And (Not VerificaTabela("AEX_PLANO")) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("BOOLCONTRATOMODARQ").AsBoolean = False
      CurrentQuery.Post
      vbTabela = False
    End If

    '		If (CurrentQuery.FieldByName("BOOLCONTLIMITCONTRARQ").AsBoolean) And (Not VerificaTabela("AEX_CONTAGEM_CONTRATO")) Then
    '			CurrentQuery.Edit
    '			CurrentQuery.FieldByName("BOOLCONTLIMITCONTRARQ").AsBoolean = False
    '			vbTabela = False
    '			CurrentQuery.Post
    '		End If



    If TABGERAARQUIVO.PageIndex = 0 Then
      ' Gera arquivo total
      CurrentQuery.Edit
      Set obj = CreateBennerObject("BSAte005.Rotinas")
      obj.GerarArquivo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      Set obj = Nothing
      CurrentQuery.FieldByName("TABGERAARQUIVO").Value = 1
      CurrentQuery.Post
      RefreshNodesWithTable("AEX_EMPCONECT")
    ElseIf TABGERAARQUIVO.PageIndex = 1 Then
      ' Gera carga nas tabelas não-parametrizáveis que estiverem marcadas
      If Not VerificaLogs Then
        CurrentQuery.Edit
      End If

      If vbTabela Then
        Set obj = CreateBennerObject("BSAte005.Rotinas")
        obj.GerarArquivo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
        Set obj = Nothing
      Else
        bsShowMessage("Algumas tabelas não podem ser exportadas por não possuirem cadastro na tabela de Tabelas Exportáveis.","i")
      End If

      If Not VerificaLogs Then
        CamposTabelasNaoParamArq
        CurrentQuery.FieldByName("TABGERAARQUIVO").Value = 1
        CurrentQuery.Post
        RefreshNodesWithTable("AEX_EMPCONECT")
      End If
    ElseIf TABGERAARQUIVO.PageIndex = 2 Then
      If Not VerificaLogs Then
        CurrentQuery.Edit
      End If

      Set obj = CreateBennerObject("BSAte005.Rotinas")
      obj.GerarArquivo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      Set obj = Nothing

      If Not VerificaLogs Then
        CamposTabelasParamArq
        CurrentQuery.FieldByName("TABGERAARQUIVO").Value = 1
        CurrentQuery.Post
        RefreshNodesWithTable("AEX_EMPCONECT")
      End If
    End If
  Else
    bsShowMessage("O registro não pode estar em edição.","i")
  End If
End Sub

Public Sub BOTAOGERACARGA_OnClick()
  Dim vbTabela As Boolean
  Dim vCargaTot As Object
  '********************* INICIO SMS - 39471 - DRUMMOND - 29/06/2005  *****************************************

  If TABCARGAS.PageIndex = 0 Then
    If CurrentQuery.State = 1 Then
      Set vCargaTot = CreateBennerObject("BSATE007.ROTINAS")
      vCargaTot.CARGATOTAL(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser)
      Set vCargaTot = Nothing

    Else
      bsShowMessage("O registro não pode estar em edição.", "i")
    End If
  ElseIf TABCARGAS.PageIndex = 1 Then
    If CurrentQuery.State = 1 Then
      If Not VerificaLogs Then
        CurrentQuery.Edit
      End If

      Set vCargaTot = CreateBennerObject("BSATE007.ROTINAS")
      vCargaTot.CARGATOTAL(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser)
      Set vCargaTot = Nothing

      ' Gera carga nas tabelas não-parametrizáveis que estiverem marcadas
      If Not VerificaLogs Then
        CamposTabelasNaoParam
        CurrentQuery.FieldByName("TABCARGAS").Value = 1
        CurrentQuery.Post
      End If
    Else
      bsShowMessage("O registro não pode estar em edição.", "i")
    End If
  ElseIf TABCARGAS.PageIndex = 2 Then
    'Tabelas parametrizáveis
    If CurrentQuery.State = 1 Then
      Set vCargaTot = CreateBennerObject("BSATE007.ROTINAS")
      vCargaTot.CARGATOTAL(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser)
      Set vCargaTot = Nothing
      If Not VerificaLogs Then
        CurrentQuery.Edit
        CamposTabelasParam
        CurrentQuery.FieldByName("TABCARGAS").Value = 1
        CurrentQuery.Post
      End If
    Else
      bsShowMessage("O registro não pode estar em edição.","i")
    End If
  End If
  RefreshNodesWithTable("AEX_EMPCONECT")
  '********************* FIM    SMS - 39471 - DRUMMOND - 28/06/2005  *****************************************
End Sub

Public Sub BOTAOGERARARQUIVO_OnClick()
  Dim obj As Object
  Set obj = CreateBennerObject("BSAte005.Rotinas")
  obj.GerarArquivo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set obj = Nothing
End Sub


Public Sub BOTAOIMPORTACAOMANUAL_OnClick()
  Dim Executar As Object
  Set Executar = CreateBennerObject("BSAte006.Rotinas")
  Executar.Importacao(CurrentSystem, "", CurrentQuery.FieldByName("NUMEROEMS").AsInteger, CurrentUser)
  Set Executar = Nothing
End Sub

Public Sub BOTAOLEITURAARQUIVOERRO_OnClick()
  Dim vLeArquivo As Object
  '********************* INICIO SMS - 39471 - DRUMMOND - 07/07/2005  *****************************************
  If CurrentQuery.State = 1 Then
    Set vLeArquivo = CreateBennerObject("BSATE007.ROTINAS")
    vLeArquivo.LerArquivo(CurrentSystem, CurrentQuery.FieldByName("CAMINHOARQUIVOERRO").AsString, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentUser)
    Set vLeArquivo = Nothing
  Else
    bsShowMessage("O registro não pode estar em edição.","i")
  End If
  '********************* FIM    SMS - 39471 - DRUMMOND - 07/07/2005  *****************************************
End Sub

Public Sub BOTAOSELARQUIVOERRO_OnClick()
  Dim Interface As Object
  Dim vPath As String


  Set Interface = CreateBennerObject("BSPRE001.Rotinas")
  vPath = Interface.SelecionarDiretorio(CurrentSystem)

  If vPath <>"" Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CAMINHOARQUIVOERRO").AsString = vPath
  End If

  Set Interface = Nothing
End Sub

Public Sub BOTAOSELCAMINHOIMP_OnClick()
  Dim Interface As Object
  Dim vPath As String


  Set Interface = CreateBennerObject("BSPRE001.Rotinas")
  vPath = Interface.SelecionarDiretorio(CurrentSystem)

  If vPath <>"" Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CAMINHOLOGRETORNO").AsString = vPath
  End If

End Sub

Public Sub BOTAOTESTAFTP_OnClick()
Dim BSAte008 As Object

Set BSAte008 = CreateBennerObject("BSAte008.Rotinas")
BSAte008.TestaFTP(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)
Set BSAte008 = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
End Sub

Public Sub TABLE_AfterPost()
  Dim Q As Object
  Dim vMacro As Variant
  Dim vCodEmp As String
  Dim vHandle As Long

  Set Q = NewQuery
  Q.Add("SELECT COUNT(1) NREC FROM Z_MACROS WHERE NOME = :NOME")
  Q.ParamByName("NOME").Value = "Importar arquivos Logs - Empresa " + vCodEmp
  Q.Active = True

  If Q.FieldByName("NREC").AsInteger > 0 Then
    Set Q = Nothing
    Exit Sub
  End If
  vCodEmp = CurrentQuery.FieldByName("NUMEROEMS").AsString
  vHandle = NewHandle("Z_MACROS")

  vMacro = "Sub Main() " + Chr(13) + _
           "  Dim Executar As Object " + Chr(13) + _
           "  Dim Q        As Object" + Chr(13) + _
           "  Set Q = NewQuery" + Chr(13) + _
           "  Q.Add(""SELECT USUARIOPADRAO, HOSTPADRAO FROM AEX_PARAMETROSGERAIS"")" + Chr(13) + _
           "  Q.Active=True " + Chr(13) + _
           "  Set Executar = CreateBennerObject(""BSAte006.Rotinas"")" + Chr(13) + _
           "  Executar.ProcessarAgendamento(CurrentSystem,""""," + vCodEmp + ", _" + Chr(13) + _
           "                                Q.FieldByName(""USUARIOPADRAO"").AsInteger, Q.FieldByName(""HOSTPADRAO"").AsString)" + Chr(13) + _
           "  Set Executar = Nothing" + Chr(13) + _
           "  Set Q = Nothing" + Chr(13) + _
           "End Sub"
  Q.Active = False

  Q.Clear
  Q.Add("INSERT INTO Z_MACROS                                                      ")
  Q.Add("  (HANDLE,NOME,MACRO,CLIDEF,TIPO)                                         ")
  Q.Add("VALUES                                                                    ")
  Q.Add("  (:HANDLE,'Importar arquivos Logs - Empresa " + vCodEmp + "',:MACRO,'N',5)   ")
  Q.ParamByName("HANDLE").Value = vHandle
  Q.ParamByName("MACRO").AsMemo = vMacro
  Q.ExecSQL


  Set Q = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If VerificaLogs Then
    CanContinue = False
    bsShowMessage("Existem processos em execução aguarde até que eles terminem.","i")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQLCodEx As Object

  '********************* INICIO SMS - 39471 - DRUMMOND - 28/06/2005  *****************************************
  Set SQLCodEx = NewQuery
  'Verifica se o código de identificação externa já existe
  SQLCodEx.Clear
  SQLCodEx.Add("SELECT COUNT(1) QTD FROM AEX_EMPCONECT WHERE NUMEROEMS = :NUMEMS AND HANDLE <> :HNDL")
  SQLCodEx.ParamByName("NUMEMS").AsInteger = CurrentQuery.FieldByName("NUMEROEMS").AsInteger
  SQLCodEx.ParamByName("HNDL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLCodEx.Active = True
  '
  If (SQLCodEx.FieldByName("QTD").AsInteger > 0) Then
    bsShowMessage("O código de identificação externa para operadora já existe.", "i")
    CanContinue = False
  End If

  Set SQLCodEx = Nothing

  '********************* FIM    SMS - 39471 - DRUMMOND - 28/06/2005  ********************************************
  CurrentQuery.FieldByName("PROCESSADO").Value = "N"
  'End If

  'SMS 87108
  'Adequação de macro para execução em ambiente Web
  If (CurrentQuery.FieldByName("TABCARGAS").AsInteger = 1) Or (CurrentQuery.FieldByName("TABCARGAS").AsInteger = 3) Then
      CamposTabelasNaoParam
  ElseIf (CurrentQuery.FieldByName("TABCARGAS").AsInteger = 1) Or (CurrentQuery.FieldByName("TABCARGAS").AsInteger = 2) Then
      CamposTabelasParam
  End If

  If (CurrentQuery.FieldByName("TABGERAARQUIVO").AsInteger = 1) Or (CurrentQuery.FieldByName("TABGERAARQUIVO").AsInteger = 3) Then
      CamposTabelasNaoParamArq
  ElseIf (CurrentQuery.FieldByName("TABGERAARQUIVO").AsInteger = 1) Or (CurrentQuery.FieldByName("TABGERAARQUIVO").AsInteger = 2) Then
      CamposTabelasParamArq
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOOLREPROCESSARARQ") Then
		BOOLREPROCESSARARQ_OnClick
	End If
	If (CommandID = "BOTAOBAIXARARQUIVOS") Then
		BOTAOBAIXARARQUIVOS_OnClick
	End If
	If (CommandID = "BOTAOCAMINHO") Then
		BOTAOCAMINHO_OnClick
	End If
	If (CommandID = "BOTAOCORRIGELOGCARGATABELAS") Then
		BOTAOCORRIGELOGCARGATABELAS_OnClick
	End If
	If (CommandID = "BOTAOGERACARGA") Then
		BOTAOGERACARGA_OnClick
	End If
	If (CommandID = "BOTAOIMPORTACAOMANUAL") Then
		BOTAOIMPORTACAOMANUAL_OnClick
	End If
	If (CommandID = "BOTAOLEITURAARQUIVOERRO") Then
		BOTAOLEITURAARQUIVOERRO_OnClick
	End If
	If (CommandID = "BOTAOSELARQUIVOERRO") Then
		BOTAOSELARQUIVOERRO_OnClick
	End If
	If (CommandID = "BOTAOSELCAMINHOIMP") Then
		BOTAOSELCAMINHOIMP_OnClick
	End If
	If (CommandID = "BOTAOTESTAFTP") Then
		BOTAOTESTAFTP_OnClick
	End If
	If (CommandID = "BOTAOGERAARQ") Then
		BOTAOGERAARQ_OnClick
	End If
	If (CommandID = "BOTAOGERARARQUIVO") Then
		BOTAOGERARARQUIVO_OnClick
	End If

End Sub
