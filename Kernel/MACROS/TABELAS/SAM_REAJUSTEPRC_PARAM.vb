'HASH: 7490F614669090B705CCBF7D8C31206B
'Macro: SAM_REAJUSTEPRC_PARAM
'#Uses "*bsShowMessage"

Option Explicit

Dim vFiltro As String

Function ChecarConsistencia(TLocal, CParamLocal, TFaixa, CParamFaixa As String)As String

  Dim SQL, SQL1 As Object

  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE FROM " + TLocal + " WHERE " + CParamLocal + " = " + CurrentQuery.FieldByName("HANDLE").AsString)
  SQL.Active = True

  If Not SQL.EOF Then

    Set SQL1 = NewQuery
    SQL1.Add("SELECT HANDLE FROM " + TFaixa + " WHERE " + CParamFaixa + " = " + SQL.FieldByName("HANDLE").AsString)
    SQL1.Active = True

    If SQL1.EOF Then
      ChecarConsistencia = "Informar dados na tabela de faixas !"
    End If

    Set SQL1 = Nothing

  End If

  Set SQL = Nothing
End Function

Public Sub ExecutaBotaoGerar(CanContinue As Boolean)

  If CurrentQuery.State = 2 Or CurrentQuery.State = 3 Then
    bsShowMessage("O registro deve ser gravado para poder gerar.", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim qTabela As Object
  Set qTabela = NewQuery
  Dim qAuxiliar As Object
  Set qAuxiliar = NewQuery

  '---Claudemir
  Dim vMsg As String

  vMsg = ChecarConsistencia("SAM_REAJUSTEPRC_DOTAC", "PARAMETRODEREAJUSTE", "SAM_REAJUSTEPRC_PARAMDOTACFX", "PARAMETRODEREAJUSTEDOTAC")
  If vMsg <>"" Then
    bsShowMessage("Dotação - Informar faixas de eventos a reajustar !", "E")
    CanContinue = False
    Exit Sub
  End If

  vMsg = ChecarConsistencia("SAM_REAJUSTEPRC_GRAU", "PARAMETRODEREAJUSTE", "SAM_REAJUSTEPRC_PARAMGRAUFX", "PARAMETRODEREAJUSTEGRAU")
  If vMsg <>"" Then
    bsShowMessage("Grau - Informar faixas de graus a reajustar !", "E")
    CanContinue = False
    Exit Sub
  End If

  vMsg = ChecarConsistencia("SAM_REAJUSTEPRC_PARAMREGIME", "PARAMETRODEREAJUSTE", "SAM_REAJUSTEPRC_PARAMREGDOTFX", "PARAMETRODEREAJUSTEDOTAC")
  If vMsg <>"" Then
    bsShowMessage("Regime de Atendimento - Informar faixas de eventos a reajustar !", "E")
    CanContinue = False
    Exit Sub
  End If

  vMsg = ChecarConsistencia("SAM_REAJUSTEPRC_PACOTE", "PARAMETROREAJUSTE", "SAM_REAJUSTEPRC_PARAMPACOTEFX", "PARAMETROSREAJUSTEPACOTE")
  If vMsg <>"" Then
    bsShowMessage("Pacote - Informar faixas de eventos a reajustar !", "E")
    CanContinue = False
    Exit Sub
  End If

  '-----------------

  qTabela.Add("SELECT NOME FROM Z_TABELAS WHERE HANDLE = :HANDLE")
  qTabela.ParamByName("HANDLE").Value = CurrentTable
  qTabela.Active = True
  If qTabela.FieldByName("NOME").AsString = "SAM_REAJUSTEPRC_PARAM" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT ASSOCIACAO, ESTADO, MUNICIPIO, PRESTADOR, REDERESTRITA FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    qAuxiliar.Active = True
  End If
  If qTabela.FieldByName("NOME").AsString = "SAM_REAJUSTEPRC_PARAMTIPO" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT ASSOCIACAO, ESTADO, MUNICIPIO, PRESTADOR, REDERESTRITA FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("REAJUSTEPRCPARAM").AsInteger
    qAuxiliar.Active = True
  End If


  'controle de acesso por municipio,estado ou por prestador
  Dim vLocal As String
  Dim Msg As String
  Dim qMunicipio As Object
  Set qMunicipio = NewQuery
  qMunicipio.Active = False
  If Not qAuxiliar.FieldByName("ASSOCIACAO").IsNull Then
    If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  If Not qAuxiliar.FieldByName("ESTADO").IsNull Then
    qMunicipio.Clear
    qMunicipio.Add("SELECT HANDLE LOCAL FROM ESTADOS WHERE HANDLE = :HANDLE")
    qMunicipio.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("ESTADO").AsInteger
    vLocal = "E"
    qMunicipio.Active = True
    vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "I", True)
    If vFiltro = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode executar essa operação", "E")
      Set qMunicipio = Nothing
      CanContinue = False
      Exit Sub
    End If
  End If
  If Not qAuxiliar.FieldByName("MUNICIPIO").IsNull Then
    qMunicipio.Clear
    qMunicipio.Add("SELECT HANDLE LOCAL FROM MUNICIPIOS WHERE HANDLE = :HANDLE")
    qMunicipio.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("MUNICIPIO").AsInteger
    vLocal = "M"
    qMunicipio.Active = True
    vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "I", True)
    If vFiltro = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode executar essa operação", "E")
      Set qMunicipio = Nothing
      CanContinue = False
      Exit Sub
    End If
  End If
  If Not qAuxiliar.FieldByName("PRESTADOR").IsNull Then
    If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  '  If Not qAuxiliar.FieldByName("REDERESTRITA").IsNull Then
  '    Exit Sub
  '  End If
  Set qMunicipio = Nothing
  'FIM controle de acesso por municipio,estado ou por prestador

  Dim Reajuste As Object
  If CurrentQuery.IsGrid Then
    bsShowMessage("Voltar à página principal para proceder com a geração.", "I")
  Else
    If CurrentQuery.State <>1 Then
      bsShowMessage("Confirme ou cancele as alterações antes de proceder com a geração.", "I")
    Else
      If Not CurrentQuery.FieldByName("DATADOPROCESSO").IsNull Then
        bsShowMessage("Processo concluído. Geração cancelada!", "I")
      Else
        If VisibleMode Then
           Set Reajuste = CreateBennerObject("BSINTERFACE0012.Rotinas")
           Reajuste.Gerar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
          Set Reajuste = Nothing
        Else
           Dim vsMensagemErro As String
           Dim viRetorno As Long

		   If CurrentQuery.FieldByName("SITUACAOGERAR").AsInteger <> 1 Then
		     Dim sqlPermiteRegerar As BPesquisa
	         Set sqlPermiteRegerar = NewQuery

	         sqlPermiteRegerar.Add("UPDATE SAM_REAJUSTEPRC_PARAM SET SITUACAOGERAR=:SITUACAOGERAR WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
	         sqlPermiteRegerar.ParamByName("SITUACAOGERAR").AsString = "1"
	         sqlPermiteRegerar.ExecSQL
	   	   End If

           Set Reajuste = CreateBennerObject("BSServerExec.ProcessosServidor")
           viRetorno = Reajuste.ExecucaoImediata(CurrentSystem, _
                                    "BSPRE011", _
                                    "RotinaReajuste_Gerar", _
                                    "Rotina de Reajuste de Preço (Geração)", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_REAJUSTEPRC_PARAM", _
                                    "SITUACAOGERAR", _
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
    End If
  End If
  If VisibleMode Then
  	SelectNode(CurrentQuery.FieldByName("handle").AsInteger,True,False)
  End If

End Sub

Public Sub BOTAOGERAR_OnClick()
	If (CurrentQuery.FieldByName("TIPOROTINA").AsInteger = 2) Then
		Dim Reajuste As Object

		Set Reajuste = CreateBennerObject("BSINTERFACE0012.Rotinas")
		Reajuste.Importar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
		Set Reajuste = Nothing
	Else
		ExecutaBotaoGerar(True)
	End If

	RefreshNodesWithTable("SAM_REAJUSTEPRC_PARAM")
End Sub

Public Sub ExecutaBotaoProcessar(CanContinue As Boolean)
  Dim qTabela As Object
  Set qTabela = NewQuery
  Dim qAuxiliar As Object
  Set qAuxiliar = NewQuery

  qTabela.Add("SELECT NOME FROM Z_TABELAS WHERE HANDLE = :HANDLE")
  qTabela.ParamByName("HANDLE").Value = CurrentTable
  qTabela.Active = True
  If qTabela.FieldByName("NOME").AsString = "SAM_REAJUSTEPRC_PARAM" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT ASSOCIACAO, ESTADO, MUNICIPIO, PRESTADOR, REDERESTRITA FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    qAuxiliar.Active = True
  End If
  If qTabela.FieldByName("NOME").AsString = "SAM_REAJUSTEPRC_PARAMTIPO" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT ASSOCIACAO, ESTADO, MUNICIPIO, PRESTADOR, REDERESTRITA FROM SAM_REAJUSTEPRC_PARAM WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("REAJUSTEPRCPARAM").AsInteger
    qAuxiliar.Active = True
  End If
  'verificando se o reajuste está gerado
  If ((CurrentQuery.FieldByName("TIPOROTINA").AsInteger = 1 And CurrentQuery.FieldByName("SITUACAOGERAR").AsInteger <> 5) Or _
  	  (CurrentQuery.FieldByName("TIPOROTINA").AsInteger = 2 And CurrentQuery.FieldByName("SITUACAOIMPORTAR").AsInteger <> 5)) Then
	bsShowMessage("Processamento não permitido. É necessária a geração dos eventos candidatos a reajuste antes de processar o mesmo.", "E")
	CanContinue = False
	Exit Sub
  End If
  'controle de acesso por municipio,estado ou por prestador
  Dim vLocal As String
  Dim Msg As String
  Dim qMunicipio As Object
  Set qMunicipio = NewQuery
  qMunicipio.Active = False
  If Not qAuxiliar.FieldByName("ASSOCIACAO").IsNull Then
    If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  If Not qAuxiliar.FieldByName("ESTADO").IsNull Then
    qMunicipio.Clear
    qMunicipio.Add("SELECT HANDLE LOCAL FROM ESTADOS WHERE HANDLE = :HANDLE")
    qMunicipio.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("ESTADO").AsInteger
    vLocal = "E"
    qMunicipio.Active = True
    vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "A", True)
    If vFiltro = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode executar essa operação", "E")
      Set qMunicipio = Nothing
      CanContinue = False
      Exit Sub
    End If
  End If
  If Not qAuxiliar.FieldByName("MUNICIPIO").IsNull Then
    qMunicipio.Clear
    qMunicipio.Add("SELECT HANDLE LOCAL FROM MUNICIPIOS WHERE HANDLE = :HANDLE")
    qMunicipio.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("MUNICIPIO").AsInteger
    vLocal = "M"
    qMunicipio.Active = True
    vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "A", True)
    If vFiltro = "N" Then
      bsShowMessage("Permissão negada. Usuário não pode executar essa operação", "E")
      Set qMunicipio = Nothing
      CanContinue = False
      Exit Sub
    End If
  End If
  If Not qAuxiliar.FieldByName("PRESTADOR").IsNull Then
    If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  '  If Not qAuxiliar.FieldByName("REDERESTRITA").IsNull Then
  '    Exit Sub
  '  End If
  Set qMunicipio = Nothing
  ' FIM controle de acesso por municipio,estado ou por prestador

  Dim Reajuste As Object
  Dim vOrigemReajuste As Integer

  If CurrentQuery.IsGrid Then
    bsShowMessage("Voltar à página principal para PROCESSAR.", "I")
  Else
    If CurrentQuery.State <>1 Then
      bsShowMessage("Confirme ou cancele as alterações antes de PROCESSAR.", "I")
    Else
      If Not CurrentQuery.FieldByName("DATADOPROCESSO").IsNull Then
        bsShowMessage("Processo concluído. Processamento cancelado!", "I")
      Else
        If VisibleMode Then
           vOrigemReajuste = NodeInternalCode
           Set Reajuste = CreateBennerObject("BSINTERFACE0012.Rotinas")
           Reajuste.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vOrigemReajuste)
          Set Reajuste = Nothing
        Else
           Dim vsMensagemErro As String
           Dim viRetorno As Long
           Dim vcContainer As CSDContainer
           Set vcContainer = NewContainer

           vcContainer.AddFields("HANDLE:INTEGER;VORIGEMREAJUSTE:INTEGER")
           vcContainer.Insert
           vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
           'Reajuste por Associação
           If WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" Or _
              WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_755" Then
             vcContainer.Field("VORIGEMREAJUSTE").AsInteger = 1
           'Reajuste por Estado
           ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" Or _
                  WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_753" Then
             vcContainer.Field("VORIGEMREAJUSTE").AsInteger = 2
           'Reajuste por Municipio
           ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702" Or _
                  WebVisionCode ="V_SAM_REAJUSTEPRC_PARAM_754" Then
             vcContainer.Field("VORIGEMREAJUSTE").AsInteger = 3
           'Reajuste por Prestador
           ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" Or _
                  WebVisionCode ="V_SAM_REAJUSTEPRC_PARAM_756" Then
             vcContainer.Field("VORIGEMREAJUSTE").AsInteger = 4
           'Reajuste por Rede Restrita
           ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" Or _
                  WebVisionCode ="V_SAM_REAJUSTEPRC_PARAM_1492" Then
             vcContainer.Field("VORIGEMREAJUSTE").AsInteger = 5

		   ElseIf CurrentEntity.TransitoryVars("WEBORIGEMREAJUSTE").IsPresent Then
			 vcContainer.Field("VORIGEMREAJUSTE").AsInteger = CurrentEntity.TransitoryVars("WEBORIGEMREAJUSTE").AsInteger
           End If

           Set Reajuste = CreateBennerObject("BSServerExec.ProcessosServidor")
           viRetorno = Reajuste.ExecucaoImediata(CurrentSystem, _
                                    "BSPRE011", _
                                    "RotinaReajuste_Processar", _
                                    "Rotina de Reajuste de Preço (Processamento)", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_REAJUSTEPRC_PARAM", _
                                    "SITUACAOPROCESSAR", _
                                    "SITUACAOGERAR", _
                                    "Geração não foi processada", _
                                    "P", _
                                    False, _
                                    vsMensagemErro, _
                                    vcContainer)

          If viRetorno = 0 Then
             bsShowMessage("Processo enviado para execução no servidor!", "I")
          Else
             bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
          End If

        End If
      End If
    End If
  End If
  If VisibleMode Then
  	SelectNode(CurrentQuery.FieldByName("handle").AsInteger,True,False)
  End If
End Sub


Public Sub BOTAOPROCESSAR_OnClick()
	ExecutaBotaoProcessar(True)
	RefreshNodesWithTable("SAM_REAJUSTEPRC_PARAM")
End Sub



Public Sub ASSOCIACAO_OnPopup(ShowPopup As Boolean)
  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT FILIAL FROM SAM_REAJUSTEPRC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC")
  SQL.Active = True


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.Z_NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"

  vCriterio = "SAM_PRESTADOR.ASSOCIACAO = 'S' "

  If vFiltro <>"" And vFiltro <>"N" Then
    vCriterio = vCriterio + " AND SAM_PRESTADOR.MUNICIPIOPAGAMENTO In (" + vFiltro + ") "
  End If


  If Not SQL.FieldByName("FILIAL").IsNull Then
    vCriterio = vCriterio + " AND SAM_PRESTADOR.FILIALPADRAO = " + SQL.FieldByName("FILIAL").AsString
  End If


  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("ASSOCIACAO").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub ESTADO_OnPopup(ShowPopup As Boolean)
  Dim vUsuario As String
  Dim VCondicao As String


  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT FILIAL FROM SAM_REAJUSTEPRC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC")
  SQL.Active = True


  vUsuario = Str(CurrentUser)
  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If
  If(VisibleMode And NodeInternalCode = 2) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" ) Then 'por estado
    VCondicao = "HANDLE IN  " + "(SELECT M.ESTADO " + _
                "       FROM Z_GRUPOUSUARIOS_FILIAIS GF, " + _
                "            MUNICIPIOS M, " + _
                "            SAM_REGIAO R " + _
                "      WHERE GF.USUARIO = " + vUsuario + _
                "        AND M.REGIAO = R.HANDLE " + _
                "        AND GF.FILIAL = R.FILIAL " + _
                "        AND GF.ALTERAR = 'S' "
    If Not SQL.FieldByName("FILIAL").IsNull Then
      VCondicao = VCondicao + "        AND M.ESTADO IN (SELECT X.ESTADO          "
      VCondicao = VCondicao + "                           FROM FILIAIS_ESTADOS X "
      VCondicao = VCondicao + "                          WHERE X.FILIAL = " + SQL.FieldByName("FILIAL").AsString
      VCondicao = VCondicao + "                        ) "
    End If
    VCondicao = VCondicao + "   ) "
  End If
  If (VisibleMode And NodeInternalCode = 3) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702") Then
    If CurrentQuery.State = 2 Then 'por municipio
      VCondicao = "HANDLE IN " + "(SELECT M.ESTADO " + _
                  "       FROM Z_GRUPOUSUARIOS_FILIAIS GF, " + _
                  "            MUNICIPIOS M, " + _
                  "            SAM_REGIAO R " + _
                  "      WHERE GF.USUARIO = " + vUsuario + _
                  "        AND M.REGIAO = R.HANDLE " + _
                  "        AND GF.FILIAL = R.FILIAL " + _
                  "        AND GF.ALTERAR = 'S' "
      If Not SQL.FieldByName("FILIAL").IsNull Then
        VCondicao = VCondicao + "        AND M.ESTADO IN (SELECT X.ESTADO          "
        VCondicao = VCondicao + "                           FROM FILIAIS_ESTADOS X "
        VCondicao = VCondicao + "                          WHERE X.FILIAL = " + SQL.FieldByName("FILIAL").AsString
        VCondicao = VCondicao + "                        ) "
      End If
      VCondicao = VCondicao + "   ) "

    End If
    If CurrentQuery.State = 3 Then
      VCondicao = "HANDLE IN " + "(SELECT M.ESTADO " + _
                  "       FROM Z_GRUPOUSUARIOS_FILIAIS GF, " + _
                  "            MUNICIPIOS M, " + _
                  "            SAM_REGIAO R " + _
                  "      WHERE GF.USUARIO = " + vUsuario + _
                  "        AND M.REGIAO = R.HANDLE " + _
                  "        AND GF.FILIAL = R.FILIAL " + _
                  "        AND GF.INCLUIR = 'S' "
      If Not SQL.FieldByName("FILIAL").IsNull Then
        VCondicao = VCondicao + "        AND M.ESTADO IN (SELECT X.ESTADO          "
        VCondicao = VCondicao + "                           FROM FILIAIS_ESTADOS X "
        VCondicao = VCondicao + "                          WHERE X.FILIAL = " + SQL.FieldByName("FILIAL").AsString
        VCondicao = VCondicao + "                        ) "
      End If
      VCondicao = VCondicao + "   ) "
    End If
  End If
  Set SQL = Nothing
  ESTADO.LocalWhere = VCondicao
End Sub

Public Sub MUNICIPIO_OnPopup(ShowPopup As Boolean)
  Dim vUsuario As String
  Dim VCondicao As String
  vUsuario = Str(CurrentUser)

  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT FILIAL FROM SAM_REAJUSTEPRC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC")
  SQL.Active = True


  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If
  '  UpdateLastUpdate("ESTADOS")
  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If
  VCondicao = "HANDLE IN " + "(SELECT M.HANDLE " + _
              "       FROM Z_GRUPOUSUARIOS_FILIAIS GF, " + _
              "            MUNICIPIOS M, " + _
              "            SAM_REGIAO R " + _
              "      WHERE GF.USUARIO = " + vUsuario + _
              "        AND M.REGIAO = R.HANDLE " + _
              "        AND GF.FILIAL = R.FILIAL " + _
              "        AND GF.ALTERAR = 'S' "
  If Not SQL.FieldByName("FILIAL").IsNull Then
    VCondicao = VCondicao + "        AND M.HANDLE IN (SELECT M1.HANDLE     "
    VCondicao = VCondicao + "                           FROM MUNICIPIOS M1 "
    VCondicao = VCondicao + "                          WHERE M1.REGIAO IN (SELECT R.HANDLE     "
    VCondicao = VCondicao + "                                                FROM SAM_REGIAO R "
    VCondicao = VCondicao + "                                               WHERE R.FILIAL = " + SQL.FieldByName("FILIAL").AsString
    VCondicao = VCondicao + "                                             ) "
    VCondicao = VCondicao + "                        ) "
  End If
  VCondicao = VCondicao + "    )"

  Set SQL = Nothing
  MUNICIPIO.LocalWhere = VCondicao
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  If CurrentQuery.State = 1 Then
    TABLE_BeforeEdit(ShowPopup)
    If ShowPopup = False Then
      Exit Sub
    End If
  End If

  Dim ProcuraDLL As Variant
  Dim vColunas As String
  Dim vCampos As String
  Dim vCriterio As String
  Dim vHandle As Long
  Dim vUsuario As String

  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Add("SELECT FILIAL FROM SAM_REAJUSTEPRC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC")
  SQL.Active = True


  vUsuario = Str(CurrentUser)
  Set ProcuraDLL = CreateBennerObject("PROCURA.PROCURAR")

  vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
  vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"


  vCriterio = "SAM_PRESTADOR.ASSOCIACAO = 'N' "

  If vFiltro <>"" And vFiltro <>"N" Then
    vCriterio = vCriterio + " AND SAM_PRESTADOR.MUNICIPIOPAGAMENTO In (" + vFiltro + ") "
  End If


  If Not SQL.FieldByName("FILIAL").IsNull Then
    vCriterio = vCriterio + " AND SAM_PRESTADOR.FILIALPADRAO = " + SQL.FieldByName("FILIAL").AsString
  End If


  vCampos = "CPF/CGC|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
  vHandle = ProcuraDLL.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", False, "")
  ShowPopup = False
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  ShowPopup = False
  Set ProcuraDLL = Nothing

End Sub

Public Sub Excluir(T As String, Handle As Long)
  Dim DEL As Object
  Set DEL = NewQuery

  DEL.Active = False
  DEL.Clear
  DEL.Add("DELETE FROM SAM_REAJUSTEPRC_" + T + "_DOT WHERE REAJUSTEPRCPARAM =:REAJUSTEPRCPARAM")
  DEL.ParamByName("REAJUSTEPRCPARAM").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEL.ExecSQL

  DEL.Active = False
  DEL.Clear
  DEL.Add("DELETE FROM SAM_REAJUSTEPRC_" + T + "_REG WHERE REAJUSTEPRCPARAM =:REAJUSTEPRCPARAM")
  DEL.ParamByName("REAJUSTEPRCPARAM").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEL.ExecSQL

  DEL.Active = False
  DEL.Clear
  DEL.Add("DELETE FROM SAM_REAJUSTEPRC_" + T + "_AN WHERE REAJUSTEPRCPARAM =:REAJUSTEPRCPARAM")
  DEL.ParamByName("REAJUSTEPRCPARAM").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEL.ExecSQL

  DEL.Active = False
  DEL.Clear
  DEL.Add("DELETE FROM SAM_REAJUSTEPRC_" + T + "_SL WHERE REAJUSTEPRCPARAM =:REAJUSTEPRCPARAM")
  DEL.ParamByName("REAJUSTEPRCPARAM").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEL.ExecSQL

  Set DEL = Nothing
End Sub

Public Sub TABLE_AfterScroll()

  Dim SQL As Object
  Dim vAux As String
  Dim vAux2 As String

  Dim vUsuario As String
  Dim sqlEstado As String
  Dim sqlMunicipio As String

  Set SQL = NewQuery
  SQL.Add("SELECT FILIAL FROM SAM_REAJUSTEPRC WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_REAJUSTEPRC")
  SQL.Active = True

  If WebMode Then
  	If WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702" Then
  		PRESTADOR.ReadOnly = True
  		ASSOCIACAO.ReadOnly = True
  		REDERESTRITA.ReadOnly = True
  	ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" Then
		PRESTADOR.ReadOnly = True
  		ASSOCIACAO.ReadOnly = True
  		REDERESTRITA.ReadOnly = True
  		MUNICIPIO.ReadOnly = True
  	ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" Then
  		PRESTADOR.ReadOnly = True
  		REDERESTRITA.ReadOnly = True
  		MUNICIPIO.ReadOnly = True
  		ESTADO.ReadOnly = True
	ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" Then
		REDERESTRITA.ReadOnly = True
  		MUNICIPIO.ReadOnly = True
  		ESTADO.ReadOnly = True
  		ASSOCIACAO.ReadOnly = True
  	ElseIf WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" Then
		MUNICIPIO.ReadOnly = True
  		ESTADO.ReadOnly = True
  		ASSOCIACAO.ReadOnly = True
  		PRESTADOR.ReadOnly = True
  	End If

	vAux =  "SAM_PRESTADOR.ASSOCIACAO = 'S' "
	vAux2 = "SAM_PRESTADOR.ASSOCIACAO = 'N' "

	vUsuario = Str(CurrentUser)

	sqlEstado = "A.HANDLE IN (SELECT M.ESTADO " + _
                "       FROM Z_GRUPOUSUARIOS_FILIAIS G, " + _
                "            MUNICIPIOS M, " + _
                "            SAM_REGIAO R " + _
                "      WHERE G.USUARIO = " + vUsuario + _
                "        AND M.REGIAO = R.HANDLE " + _
                "        AND G.FILIAL = R.FILIAL "
	sqlMunicipio = 	 "A.HANDLE IN (SELECT M.HANDLE " + _
					 "       FROM Z_GRUPOUSUARIOS_FILIAIS G, " + _
					 "            MUNICIPIOS M, " + _
					 "            SAM_REGIAO R " + _
					 "      WHERE G.USUARIO = " + vUsuario + _
					 "        AND M.REGIAO = R.HANDLE " + _
					 "        AND G.FILIAL = R.FILIAL "



	If CurrentQuery.State = 3 Then
    	sqlEstado = sqlEstado + " AND G.INCLUIR = 'S' "
    	sqlMunicipio = sqlMunicipio + " AND G.INCLUIR = 'S' "
    Else
		sqlEstado = sqlEstado + " AND G.ALTERAR = 'S' "
		sqlMunicipio = sqlMunicipio + " AND G.ALTERAR = 'S' "
    End If

  	If Not SQL.FieldByName("FILIAL").IsNull Then
	    ASSOCIACAO.WebLocalWhere = vAux + " AND SAM_PRESTADOR.FILIALPADRAO = " + SQL.FieldByName("FILIAL").AsString
	    PRESTADOR.WebLocalWhere = vAux2 + " AND SAM_PRESTADOR.FILIALPADRAO = " + SQL.FieldByName("FILIAL").AsString
	    sqlEstado = sqlEstado + " AND M.ESTADO IN (SELECT X.ESTADO          " + _
	    						"					 FROM FILIAIS_ESTADOS X " + _
								"                          WHERE X.FILIAL = " + SQL.FieldByName("FILIAL").AsString + _
								"                 )	                        "

		sqlMunicipio = sqlMunicipio +   " AND M.HANDLE IN (SELECT M1.HANDLE                             " + _
										"                  FROM MUNICIPIOS M1                           " + _
										"                        WHERE M1.REGIAO IN (SELECT R.HANDLE    " + _
										"                                             FROM SAM_REGIAO R " + _
										"                                             WHERE R.FILIAL =  " + SQL.FieldByName("FILIAL").AsString + _
										"                                           )                   " + _
										"                  )                                            "

	End If

	ESTADO.WebLocalWhere = sqlEstado + ")"
	MUNICIPIO.WebLocalWhere = sqlMunicipio + ")"

  End If

  If (CurrentQuery.FieldByName("TIPOROTINA").AsInteger = 2) Then
	BOTAOGERAR.Caption = "Importar"
	SITUACAOIMPORTAR.Visible = True
	SITUACAOGERAR.Visible = False
  Else
	BOTAOGERAR.Caption = "Gerar"
	SITUACAOIMPORTAR.Visible = False
	SITUACAOGERAR.Visible = True
  End If

  ARQUIVODOTACAO.ReadOnly = False
  If Not CurrentQuery.FieldByName("SITUACAOIMPORTAR").IsNull Then
  	ARQUIVODOTACAO.ReadOnly = (CurrentQuery.FieldByName("SITUACAOIMPORTAR").AsInteger > 2)
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vLocal As String
  Dim LOCAL As Long
  Dim qPermissao As Object
  Set qPermissao = NewQuery
  Dim Msg As String

    'CARGA -ASSOCIASSOES
    If (VisibleMode And NodeInternalCode =  1) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" ) Then

      'se estiver alterando
      If CurrentQuery.State = 2 Then
        'vFiltro =checkPermissaoFilial(CurrentSystem,"A","P",Msg)
        'qPermissao.Active =False
        'qPermissao.Clear
        'qPermissao.Add("SELECT DISTINCT P.HANDLE ")
        'qPermissao.Add("  FROM SAM_PRESTADOR P, ")
        'qPermissao.Add("       SAM_REGIAO R, ")
        'qPermissao.Add("       MUNICIPIOS M ")
        'qPermissao.Add(" WHERE P.HANDLE = :HANDLE ")
        'qPermissao.Add("   AND P.FILIALPADRAO = R.FILIAL ")
        'qPermissao.Add("   AND M.REGIAO = R.HANDLE ")
        'qPermissao.Add("   AND M.HANDLE IN " +vFiltro)
        'qPermissao.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("ASSOCIACAO").AsInteger
        'qPermissao.Active =True
        'qPermissao.First
        'If qPermissao.EOF Then
        If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
          bsShowMessage("Permissão negada! Usuário não pode alterar.", "E")
          CanContinue = False
          Set qPermissao = Nothing
          Exit Sub
        End If
        Set qPermissao = Nothing
      End If
      'se estiver inserindo
      If CurrentQuery.State = 3 Then
        'vFiltro =checkPermissaoFilial(CurrentSystem,"I","P",Msg)
        'qPermissao.Active =False
        'qPermissao.Clear
        'qPermissao.Add("SELECT DISTINCT P.HANDLE ")
        'qPermissao.Add("  FROM SAM_PRESTADOR P, ")
        'qPermissao.Add("       SAM_REGIAO R, ")
        'qPermissao.Add("       MUNICIPIOS M ")
        'qPermissao.Add(" WHERE P.HANDLE = :HANDLE ")
        'qPermissao.Add("   AND P.FILIALPADRAO = R.FILIAL ")
        'qPermissao.Add("   AND M.REGIAO = R.HANDLE ")
        'qPermissao.Add("   AND M.HANDLE IN " +vFiltro)
        'qPermissao.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("ASSOCIACAO").AsInteger
        'qPermissao.Active =True
        'qPermissao.First
        'If qPermissao.EOF Then
        If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
         bsShowMessage("Permissão negada! Usuário não pode incluir.", "E")
          CanContinue = False
          Set qPermissao = Nothing
          Exit Sub
        End If
        Set qPermissao = Nothing
      End If

 	End If
      'CARGA -ESTADOS
    If (VisibleMode And NodeInternalCode =  2) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" ) Then
      vLocal = "E"
      LOCAL = CurrentQuery.FieldByName("ESTADO").AsInteger
      If CurrentQuery.State = 2 And vLocal <>"nada" Then
        If checkPermissao(CurrentSystem, CurrentUser, vLocal, LOCAL, "A") = "N" Then
          bsShowMessage("Permissão negada! Usuario não pode alterar", "E")
          CanContinue = False
          Exit Sub
        End If
      End If
      If CurrentQuery.State = 3 And vLocal <>"nada" Then
        If checkPermissao(CurrentSystem, CurrentUser, vLocal, LOCAL, "I") = "N" Then
          bsShowMessage("Permissão negada! Usuario não pode incluir", "E")
          CanContinue = False
          Exit Sub
        End If
      End If
	End If


      'CARGA -MUNICIPIOS
    If (VisibleMode And NodeInternalCode =  3)  Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702") Then
      vLocal = "M"
      LOCAL = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
      If CurrentQuery.State = 2 And vLocal <>"nada" Then
        If checkPermissao(CurrentSystem, CurrentUser, vLocal, LOCAL, "A") = "N" Then
          bsShowMessage("Permissão negada! Usuario não pode alterar", "E")
          CanContinue = False
          Exit Sub
        End If
      End If
      If CurrentQuery.State = 3 And vLocal <>"nada" Then
        If checkPermissao(CurrentSystem, CurrentUser, vLocal, LOCAL, "I") = "N" Then
          bsShowMessage("Permissão negada! Usuario não pode incluir", "E")
          CanContinue = False
          Exit Sub
        End If
      End If

    End If

      'CARGA -PRESTADORES
    If (VisibleMode And NodeInternalCode = 4) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" ) Then
      'se estiver alterando
      If CurrentQuery.State = 2 Then
        'vFiltro =checkPermissaoFilial(CurrentSystem,"A","P",Msg)
        'qPermissao.Active =False
        'qPermissao.Clear
        'qPermissao.Add("SELECT DISTINCT P.HANDLE ")
        'qPermissao.Add("  FROM SAM_PRESTADOR P, ")
        'qPermissao.Add("       SAM_REGIAO R, ")
        'qPermissao.Add("       MUNICIPIOS M ")
        'qPermissao.Add(" WHERE P.HANDLE = :HANDLE ")
        'qPermissao.Add("   AND P.FILIALPADRAO = R.FILIAL ")
        'qPermissao.Add("   AND M.REGIAO = R.HANDLE ")
        'qPermissao.Add("   AND M.HANDLE IN " +vFiltro)
        'qPermissao.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("PRESTADOR").AsInteger
        'qPermissao.Active =True
        'qPermissao.First
        'If qPermissao.EOF Then
        If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
          bsShowMessage("Permissão negada! Usuário não pode alterar.", "E")
          CanContinue = False
          Set qPermissao = Nothing
          Exit Sub
        End If
        Set qPermissao = Nothing
      End If
      'se estiver inserindo
      If CurrentQuery.State = 3 Then
        'vFiltro =checkPermissaoFilial(CurrentSystem,"I","P",Msg)
        'qPermissao.Active =False
        'qPermissao.Clear
        'qPermissao.Add("SELECT DISTINCT P.HANDLE ")
        'qPermissao.Add("  FROM SAM_PRESTADOR P, ")
        'qPermissao.Add("       SAM_REGIAO R, ")
        'qPermissao.Add("       MUNICIPIOS M ")
        'qPermissao.Add(" WHERE P.HANDLE = :HANDLE ")
        'qPermissao.Add("   AND P.FILIALPADRAO = R.FILIAL ")
        'qPermissao.Add("   AND M.REGIAO = R.HANDLE ")
        'qPermissao.Add("   AND M.HANDLE IN " +vFiltro)
        'qPermissao.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("PRESTADOR").AsInteger
        'qPermissao.Active =True
        'qPermissao.First
        'If qPermissao.EOF Then
        If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
          bsShowMessage("Permissão negada! Usuário não pode incluir.", "E")
          CanContinue = False
          Set qPermissao = Nothing
          Exit Sub
        End If
        Set qPermissao = Nothing
      End If

	End If

      'CARGA -REDE RESTRITA
    If (VisibleMode And NodeInternalCode = 5) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" ) Then
      vLocal = "nada"
	End If

    If (VisibleMode And NodeInternalCode =  1) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" ) Then
      CanContinue = VerAssociacao
    End If
    If (VisibleMode And NodeInternalCode = 2) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" )Then
      CanContinue = VerEstado
    End If
    If (VisibleMode And NodeInternalCode =  3)  Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702") Then
      CanContinue = VerMunicipio
    End If
    If (VisibleMode And NodeInternalCode =  4) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" ) Then
      CanContinue = VerPrestador
    End If
    If (VisibleMode And NodeInternalCode =  5) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" ) Then
      CanContinue = VerRedeRestrita
	End If
  If CanContinue = False Then
    bsShowMessage("Os Campos de seleção não podem ser nulos", "E")
    'CurrentQuery.Cancel
  End If
End Sub

Public Function VerAssociacao As Boolean
  VerAssociacao = True

  If CurrentQuery.FieldByName("ASSOCIACAO").IsNull Then
    VerAssociacao = False
  End If

  If( _
     (CurrentQuery.FieldByName("CLASSEA").AsString = "N" Or CurrentQuery.FieldByName("CLASSEA").IsNull)And _
     (CurrentQuery.FieldByName("CLASSEB").AsString = "N" Or CurrentQuery.FieldByName("CLASSEB").IsNull)And _
     (CurrentQuery.FieldByName("CLASSEC").AsString = "N" Or CurrentQuery.FieldByName("CLASSEC").IsNull)And _
     (CurrentQuery.FieldByName("CLASSED").AsString = "N" Or CurrentQuery.FieldByName("CLASSED").IsNull)And _
     (CurrentQuery.FieldByName("CLASSEE").AsString = "N" Or CurrentQuery.FieldByName("CLASSEE").IsNull)And _
     (CurrentQuery.FieldByName("CLASSEF").AsString = "N" Or CurrentQuery.FieldByName("CLASSEF").IsNull)And _
     (CurrentQuery.FieldByName("CLASSEG").AsString = "N" Or CurrentQuery.FieldByName("CLASSEG").IsNull)And _
     (CurrentQuery.FieldByName("CLASSEN").AsString = "N" Or CurrentQuery.FieldByName("CLASSEN").IsNull) _
     )Then
  VerAssociacao = False
End If
End Function


Public Function VerEstado As Boolean
  If CurrentQuery.FieldByName("ESTADO").IsNull Then
    VerEstado = False
  Else
    VerEstado = True
  End If
End Function


Public Function VerMunicipio As Boolean
  If CurrentQuery.FieldByName("ESTADO").IsNull Or CurrentQuery.FieldByName("MUNICIPIO").IsNull Then
    VerMunicipio = False
  Else
    VerMunicipio = True
  End If
End Function


Public Function VerPrestador As Boolean
  If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
    VerPrestador = False
  Else
    VerPrestador = True
  End If
End Function


Public Function VerRedeRestrita As Boolean
  If CurrentQuery.FieldByName("REDERESTRITA").IsNull Then
    VerRedeRestrita = False
  Else
    VerRedeRestrita = True
  End If
End Function


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim vLocal As String
  Dim Msg As String
  Dim qMunicipio As Object
  Set qMunicipio = NewQuery
  qMunicipio.Active = False
    If (VisibleMode And NodeInternalCode = 1) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" ) Or (WebMode And CurrentQuery.FieldByName("ASSOCIACAO").AsInteger > 0 ) Then
      If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then

		bsShowMessage(Msg, "E")


        CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 2) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" )  Or (WebMode And CurrentQuery.FieldByName("ESTADO").AsInteger > 0 And CurrentQuery.FieldByName("MUNICIPIO").AsInteger <= 0 ) Then
      qMunicipio.Clear
      qMunicipio.Add("SELECT HANDLE LOCAL FROM ESTADOS WHERE HANDLE = :HANDLE")
      qMunicipio.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ESTADO").AsInteger
      vLocal = "E"
      qMunicipio.Active = True
      If checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "E", True) = "N" Then

		bsShowMessage("Permissão negada. Usuário não pode Excluir", "E")

        CanContinue = False
        Set qMunicipio = Nothing
        Exit Sub
      End If
      Set qMunicipio = Nothing
    End If

    If (VisibleMode And NodeInternalCode = 3)  Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702") Or (WebMode And CurrentQuery.FieldByName("MUNICIPIO").AsInteger > 0 ) Then
      qMunicipio.Clear
      qMunicipio.Add("SELECT HANDLE LOCAL FROM MUNICIPIOS WHERE HANDLE = :HANDLE")
      qMunicipio.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
      vLocal = "M"
      qMunicipio.Active = True
      If checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "E", True) = "N" Then

		bsShowMessage("Permissão negada. Usuário não pode excluir", "E")

        CanContinue = False
        Set qMunicipio = Nothing
        Exit Sub
      End If
      Set qMunicipio = Nothing
    End If

    If (VisibleMode And NodeInternalCode = 4) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" ) Or (WebMode And CurrentQuery.FieldByName("PRESTADOR").AsInteger > 0 ) Then
      If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
			bsShowMessage(Msg, "E")

        CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 5) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" ) Then
      Exit Sub
    End If

  Dim DEL As Object
  Set DEL = NewQuery
  Excluir "PRESTADOR", CurrentQuery.FieldByName("HANDLE").AsInteger
  Excluir "MUNICIPIO", CurrentQuery.FieldByName("HANDLE").AsInteger
  Excluir "ESTADO", CurrentQuery.FieldByName("HANDLE").AsInteger
  Excluir "REDE", CurrentQuery.FieldByName("HANDLE").AsInteger

  DEL.Add("DELETE FROM SAM_REAJUSTEPRC_PARAMTIPO WHERE REAJUSTEPRCPARAM = :REAJUSTEPRCPARAM")
  DEL.ParamByName("REAJUSTEPRCPARAM").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  DEL.ExecSQL
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If WebMode Then
  	CONVENIO.WebLocalWhere = "A.HANDLE IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"
  ElseIf VisibleMode Then
  	CONVENIO.LocalWhere = "SAM_CONVENIO.HANDLE IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"
  End If


  Dim vLocal As String
  Dim Msg As String
  Dim qMunicipio As Object
  Set qMunicipio = NewQuery
  qMunicipio.Active = False

    If (VisibleMode And NodeInternalCode = 1) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" ) Or (WebMode And CurrentQuery.FieldByName("ASSOCIACAO").AsInteger > 0 ) Then
      vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)
      If vFiltro = "N" Then
		bsShowMessage(Msg, "E")
   		CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 2) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" )  Or (WebMode And CurrentQuery.FieldByName("ESTADO").AsInteger > 0 And CurrentQuery.FieldByName("MUNICIPIO").AsInteger <= 0 ) Then
      qMunicipio.Clear
      qMunicipio.Add("SELECT HANDLE LOCAL FROM ESTADOS WHERE HANDLE = :HANDLE")
      qMunicipio.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ESTADO").AsInteger
      vLocal = "E"
      qMunicipio.Active = True
      vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, CurrentQuery.FieldByName("ESTADO").AsInteger, "A", True)
      If vFiltro = "N" Then
		bsShowMessage("Permissão negada. Usuário não pode alterar", "E")

        CanContinue = False
        Set qMunicipio = Nothing
        Exit Sub
      End If
      Set qMunicipio = Nothing
    End If

    If (VisibleMode And NodeInternalCode =3) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702") Or (WebMode And CurrentQuery.FieldByName("MUNICIPIO").AsInteger > 0 ) Then
      qMunicipio.Clear
      qMunicipio.Add("SELECT HANDLE LOCAL FROM MUNICIPIOS WHERE HANDLE = :HANDLE")
      qMunicipio.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
      vLocal = "M"
      qMunicipio.Active = True
      vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, qMunicipio.FieldByName("LOCAL").AsInteger, "A", True)
      If vFiltro = "N" Then

		bsShowMessage("Permissão negada. Usuário não pode alterar", "E")

      CanContinue = False
        Set qMunicipio = Nothing
        Exit Sub
      End If
      Set qMunicipio = Nothing
    End If
    If (VisibleMode And NodeInternalCode = 4) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" )  Or (WebMode And CurrentQuery.FieldByName("PRESTADOR").AsInteger > 0 ) Then
      vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)
      If vFiltro = "N" Then

		bsShowMessage(Msg, "E")

        CanContinue = False
        Exit Sub
      End If
    End If
    If (VisibleMode And NodeInternalCode = 5) Or (WebMode And WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" ) Then
      Exit Sub
    End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  If WebMode Then
  	CONVENIO.WebLocalWhere = "A.HANDLE IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"
  ElseIf VisibleMode Then
  	CONVENIO.LocalWhere = "SAM_CONVENIO.HANDLE IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"
  End If


  Dim vLocal As String
  Dim Msg As String
  Dim WebOrigem As String

	If WebMode And CurrentEntity.TransitoryVars("WEBORIGEMREAJUSTEPRECO").IsPresent Then
			WebOrigem = CurrentEntity.TransitoryVars("WEBORIGEMREAJUSTEPRECO").AsString
	End If

    If (VisibleMode And NodeInternalCode = 1) Or (WebMode And ( WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_703" Or WebOrigem = "1" ) ) Then
      vFiltro = checkPermissaoFilial(CurrentSystem, "I", "P", Msg)
      If vFiltro = "N" Then

		bsShowMessage(Msg, "E")

        CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 2) Or (WebMode And ( WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_701" Or WebOrigem = "2") )  Then
      vLocal = "E"
      vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, 0, "I", True)
      If vFiltro = "N" Then

		bsShowMessage("Permissão negada. Usuário não pode incluir", "E")

        CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 3)  Or (WebMode And (WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_702" Or WebOrigem = "3")) Then
      vLocal = "M"
      vFiltro = checkPermissao(CurrentSystem, CurrentUser, vLocal, 0, "I", True)
      If vFiltro = "N" Then

		bsShowMessage("Permissão negada. Usuário não pode incluir", "E")

        CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 4)  Or (WebMode And (WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_704" Or WebOrigem = "4") ) Then
      vFiltro = checkPermissaoFilial(CurrentSystem, "I", "P", Msg)
      If vFiltro = "N" Then

		bsShowMessage(Msg, "E")

        CanContinue = False
        Exit Sub
      End If
    End If

    If (VisibleMode And NodeInternalCode = 5) Or (WebMode And (WebVisionCode = "V_SAM_REAJUSTEPRC_PARAM_705" Or WebOrigem = "5") )  Then
      Exit Sub
    End If
End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
  SQL.Active = True

  If SQL.FieldByName("TOTAL").AsInteger = 1 Then
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")
    SQL.Active = True
    CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("REAJUSTEPRC").AsInteger = RecordHandleOfTable("SAM_REAJUSTEPRC")
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOGERAR"
			ExecutaBotaoGerar(CanContinue)
		Case "BOTAOPROCESSAR"
			ExecutaBotaoProcessar(CanContinue)
	End Select
End Sub

Public Sub TIPOROTINA_OnChange()
	CurrentQuery.UpdateRecord

	If CurrentQuery.FieldByName("TIPOROTINA").AsInteger = 1 Then
		SITUACAOGERAR.Visible = True
		SITUACAOIMPORTAR.Visible = False
	Else
		SITUACAOGERAR.Visible = False
		SITUACAOIMPORTAR.Visible = True
	End If
End Sub
