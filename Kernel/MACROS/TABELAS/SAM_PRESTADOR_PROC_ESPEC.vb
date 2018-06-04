'HASH: EBFD842B3D76B10959A3CA65042E1AFE
'Macro: SAM_PRESTADOR_PROC_ESPEC

'#Uses "*bsShowMessage"

'Mauricio Ibelli -04/01/2002 -sms3165 -Se filial padrao do prestador for nulo não checar responsavel
'Atualizacao -13/01/2003 -Claudemir

Dim Mensagem As String
Dim vUsuario As Boolean

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Mensagem = ""

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  'O processo finalizado não pode ser alterado.

  SQL.Add("SELECT SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And((SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser)Or(SQL.FieldByName("FILIALPADRAO").IsNull)), True, False)

  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo Finalizado! Operação não permitida." + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <>CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub ESPECIALIDADE_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim ProcuraEspec As Long
  Dim SQL As Object
  Dim vPrestador As String

  If CurrentQuery.FieldByName("OPERACAO").IsNull Then
    ShowPopup = False
    MsgBox "Selecione uma operação!"
    Exit Sub
  End If

  ShowPopup = False
  ProcuraEspec = 0
  OPERACAO.ReadOnly = False
  If CurrentQuery.FieldByName("ESPECIALIDADE").IsNull Then
    If MsgBox("Após selecionar uma especialidade o campo 'Operação'  " + (Chr(13)) + _
               "será desabilitado, não podendo ser alterado! " + (Chr(13)) + _
               "Deseja Continua?", vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  Set Interface = CreateBennerObject("Procura.Procurar")

  If CurrentQuery.FieldByName("OPERACAO").Value <>1 Then
    vColunas = "SAM_ESPECIALIDADE.DESCRICAO|SAM_PRESTADOR_ESPECIALIDADE.DATAINICIAL|SAM_PRESTADOR_ESPECIALIDADE.DATAFINAL"
    vCriterio = "SAM_PRESTADOR_ESPECIALIDADE.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
    vCampos = "Especialidade|Data inicial|Datafinal"
    vTabela = "SAM_PRESTADOR_ESPECIALIDADE|SAM_ESPECIALIDADE[SAM_ESPECIALIDADE.HANDLE=SAM_PRESTADOR_ESPECIALIDADE.ESPECIALIDADE]"
    Set SQL = NewQuery
    SQL.Add("SELECT NOME FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").Value
    SQL.Active = True
    vPrestador = SQL.FieldByName("NOME").AsString
    SQL.Active = False

    ProcuraEspec = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Especialidades do prestador " + vPrestador, True, "")
  Else
    vColunas = "DESCRICAO"
    vCriterio = ""
    vCampos = "Especialidade"
    vTabela = "SAM_ESPECIALIDADE"
    ProcuraEspec = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Todas as especialidades", True, "")
    'CurrentQuery.FieldByName("DATAINICIAL").AsDateTime =ServerDate
  End If

  If ProcuraEspec <>0 Then
    If CurrentQuery.FieldByName("OPERACAO").Value <>1 Then
      SQL.Clear
      SQL.Add("SELECT E.HANDLE, P.DATAINICIAL, P.DATAFINAL                           ")
      SQL.Add("  FROM SAM_ESPECIALIDADE            E                                 ")
      SQL.Add("  JOIN SAM_PRESTADOR_ESPECIALIDADE  P ON (P.ESPECIALIDADE = E.HANDLE) ")
      SQL.Add(" WHERE P.HANDLE = :HANDLE                                               ")
      SQL.ParamByName("HANDLE").Value = ProcuraEspec
      '-------------------------------------------------------

      '-------------------------------------------------------
      SQL.Active = False
      SQL.Active = True

      CurrentQuery.FieldByName("DATAINICIAL").Value = SQL.FieldByName("DATAINICIAL").Value
      CurrentQuery.FieldByName("DATAFINAL").Value = SQL.FieldByName("DATAFINAL").Value
      CurrentQuery.FieldByName("ESPECIALIDADE").Value = SQL.FieldByName("HANDLE").Value
      CurrentQuery.FieldByName("PRESTADORESPECIALIDADE").Value = ProcuraEspec

      DATAINICIAL.ReadOnly = True
      DATAFINAL.ReadOnly = True
      If CurrentQuery.FieldByName("OPERACAO").Value = 3 Then
        DATAINICIAL1.ReadOnly = False
        DATAFINAL1.ReadOnly = False
      End If

    Else
      CurrentQuery.FieldByName("ESPECIALIDADE").Value = ProcuraEspec
    End If
    OPERACAO.ReadOnly = True
  End If


  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterCommitted()
  If CurrentQuery.FieldByName("VISUALIZARCENTRAL").AsString = "S" And _
                              CurrentQuery.FieldByName("AREALIVRO").IsNull Then
    Dim vsXMLSelecao As String
    Dim vsMensagem   As String

    vsXMLSelecao = ""

    If WebMode Then
      Dim qSQL As Object

	  Set qSQL = NewQuery

	  qSQL.Clear

	  qSQL.Add("SELECT E.HANDLE HENDERECO,              ")
	  qSQL.Add("       A.HANDLE HAREA                   ")
	  qSQL.Add("  FROM SAM_PRESTADOR_ENDERECO E,        ")
	  qSQL.Add("       SAM_DIMENSIONAMENTO    D,        ")
	  qSQL.Add("       SAM_AREALIVRO          A         ")
	  qSQL.Add(" WHERE A.HANDLE        = D.AREALIVRO    ")
	  qSQL.Add("   AND E.PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString)
	  qSQL.Add("   AND D.ESPECIALIDADE = " + CurrentQuery.FieldByName("ESPECIALIDADE").AsString)
	  qSQL.Add("   AND E.ATENDIMENTO   = 'S'            ")
	  qSQL.Add("   AND E.DATACANCELAMENTO  IS NULL      ")
	  qSQL.Active = True

      If Not qSQL.EOF Then
		Dim vcContainer As CSDContainer

        Set vcContainer = NewContainer

        vcContainer.GetFieldsFromQuery(qSQL.TQuery)
        vcContainer.LoadAllFromQuery(qSQL.TQuery)

        vsXMLSelecao = vcContainer.GetXML

        Set vcContainer = Nothing
	    Set qSQL = Nothing
      End If
    Else
	  Dim dllBSInterface0020_SelecaoEnderecoAreaLivro As Object

	  Set dllBSInterface0020_SelecaoEnderecoAreaLivro = CreateBennerObject("BSINTERFACE0020.SelecaoEnderecoAreaLivro")

	  If dllBSInterface0020_SelecaoEnderecoAreaLivro.Exec(CurrentSystem, _
			                                              CurrentQuery.FieldByName("PRESTADOR").AsInteger, _
			                                              CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
			                                              True, _
			                                              vsXMLSelecao, _
			                                              vsMensagem) = 1 Then
        bsShowMessage("Erro na seleção de endereços e área de livro: " + vsMensagem, "I")
      End If

      Set dllBSInterface0020_SelecaoEnderecoAreaLivro = Nothing
    End If

	If vsXMLSelecao <> "" Then
      Dim dllBSPre001_AtualizacaoEspecialidade As Object

      Set dllBSPre001_AtualizacaoEspecialidade = CreateBennerObject("BSPRE001.AtualizacaoEspecialidade")

      If dllBSPre001_AtualizacaoEspecialidade.Processo(CurrentSystem, _
			                                           CurrentQuery.FieldByName("HANDLE").AsInteger, _
			                                           vsXMLSelecao, _
			                                           vsMensagem) Then
        bsShowMessage("Erro na atualização do processo: " + vsMensagem, "I")
      Else
        If WebMode Then
          bsShowMessage("Especialidade incluída no 'Livro de credenciados' para todos os endereços de atendimento ativos do Prestador!", "I")
        End If
	  End If

      Set dllBSPre001_AtualizacaoEspecialidade = Nothing
    Else
      Dim qUpd As Object

      Set qUpd = NewQuery

      StartTransaction

      qUpd.Add("UPDATE SAM_PRESTADOR_PROC_ESPEC SET")
      qUpd.Add("  VISUALIZARCENTRAL = 'N'")
      qUpd.Add("WHERE HANDLE = :HPROCESSOESPEC")

      qUpd.ParamByName("HPROCESSOESPEC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

      qUpd.ExecSQL

      Commit

      If VisibleMode Then
        bsShowMessage("Opção ""Visualizar na central de atendimento"" será desmarcada!" + Chr(13) + _
                      "Motivo: Nenhum registro foi selecionado.", "I")
      Else
        bsShowMessage("Opção ""Visualizar na central de atendimento"" será desmarcada!" + Chr(13) + _
                      "Motivo: Prestador não possui endereço ou não existe área no livro para a especialidade.", "I")
      End If

      Set qUpd = Nothing
    End If
  End If
End Sub

Public Sub TABLE_AfterPost()
  'alterado para aftercommitted a pedido do larini - sms 59930
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL, SQL1 As Object

  DATAINICIAL1.ReadOnly = True
  DATAFINAL1.ReadOnly = True

  ESPECIALIDADE.ReadOnly = False
  DATAINICIAL.ReadOnly = False
  DATAFINAL.ReadOnly = False
  OPERACAO.ReadOnly = False

  If CurrentQuery.FieldByName("OPERACAO").Value <>1 Then
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
  End If

  If CurrentQuery.FieldByName("OPERACAO").Value = 3 Then
    DATAINICIAL1.ReadOnly = False
    DATAFINAL1.ReadOnly = False
  Else
  End If

  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    Set SQL = NewQuery
    SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC_GRP WHERE PRESTADORPROCESSO = :PRESTADORPROCESSO")
    SQL.ParamByName("PRESTADORPROCESSO").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.Active = True
    If Not SQL.EOF Then
      ESPECIALIDADE.ReadOnly = True
      DATAINICIAL1.ReadOnly = True
      DATAFINAL1.ReadOnly = True
      DATAINICIAL.ReadOnly = True
      DATAFINAL.ReadOnly = True
    Else
      ESPECIALIDADE.ReadOnly = False
      Set SQL1 = NewQuery
      SQL1.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESP_REDE WHERE PROCESSOESPECIALIDADE = :PROCESSOESPECIALIDADE")
      SQL1.ParamByName("PROCESSOESPECIALIDADE").Value = CurrentQuery.FieldByName("HANDLE").Value
      SQL1.Active = True
      If Not SQL1.EOF Then
        ESPECIALIDADE.ReadOnly = True
        DATAINICIAL1.ReadOnly = True
        DATAFINAL1.ReadOnly = True
        DATAINICIAL.ReadOnly = True
        DATAFINAL.ReadOnly = True
      End If
    End If
    Set SQL = Nothing
    Set SQL1 = Nothing
  End If

  If WebMode Then
    ESPECIALIDADE.WebLocalWhere = "(A.HANDLE IN (SELECT HANDLE FROM SAM_ESPECIALIDADE) AND @CAMPO(OPERACAO) = '1')               " + _
								  "       OR                                                                                     " + _
								  "       (A.HANDLE IN (SELECT E.HANDLE                                                          " + _
								  "                       FROM SAM_ESPECIALIDADE E                                               " + _
								  "                       JOIN SAM_PRESTADOR_ESPECIALIDADE EP ON (E.HANDLE = EP.ESPECIALIDADE)   " + _
								  "                      WHERE EP.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString   + _
								  "                    )                                                                         " + _
								  "        AND @CAMPO(OPERACAO) <> '1'                                                           " + _
								  "        )                                                                                     "

  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  '-----------------------------------------------------------
  If CurrentQuery.FieldByName("OPERACAO").Value = 1 Then
    Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
    Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString

    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADE", Condicao)

    If Linha <>"" Then
      CanContinue = False
      bsShowMessage(Linha, "E")
      Exit Sub
    End If
    Set Interface = Nothing
  Else
    If CurrentQuery.FieldByName("OPERACAO").Value = 3 Then
      If Not CurrentQuery.FieldByName("DATAINICIAL1").IsNull Then
        Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
        Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString
        Condicao = Condicao + " AND HANDLE <> " + CurrentQuery.FieldByName("PRESTADORESPECIALIDADE").AsString
        Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime, CurrentQuery.FieldByName("DATAFINAL1").AsDateTime, "ESPECIALIDADE", Condicao)
        If Linha <>"" Then
          CanContinue = False
          bsShowMessage(Linha, "E")
          Exit Sub
        End If
        Set Interface = Nothing
      End If
    End If
  End If
  '-----------------------------------------------------------


  If CurrentQuery.FieldByName("OPERACAO").Value <>3 Then
    CurrentQuery.FieldByName("DATAINICIAL1").Clear
    CurrentQuery.FieldByName("DATAFINAL1").Clear
  End If

  SQL.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADE A WHERE A.PRESTADOR = :PREST AND A.ESPECIALIDADE = :ESPEC")
  SQL.ParamByName("PREST").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  SQL.ParamByName("ESPEC").Value = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  SQL.Active = True
  If SQL.EOF Then
    If CurrentQuery.FieldByName("OPERACAO").AsString = "2" Then
      CanContinue = False
      bsShowMessage("Especialidade não encontrada. Operação incoerente!", "E")
    End If
  Else
    If CurrentQuery.FieldByName("OPERACAO").AsString = "1" Then
      'Cancontinue =False
      'MsgBox "Especialidade já cadastrada. Operação de Alteração!"
    ElseIf CurrentQuery.FieldByName("OPERACAO").AsString = "2" Then 'verificar se a exclusão é da especialidade principal
      If SQL.FieldByName("PRINCIPAL").AsString = "S" Then
        'A especialidade principal deixou de ser exigida sendo assim não é mais preciso marcar uma comp principal
        'SQL.Clear
        'SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC WHERE PRESTADORPROCESSO = :PRESTPROC AND PRINCIPAL = 'S'")
        'SQL.Add("AND OPERACAO = 'I'")
        'SQL.ParamByName("PRESTPROC").Value =CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
        'SQL.Active =True
        'If SQL.EOF Then
        '  Cancontinue =False
        '  MsgBox "Não é permitido excluir a especialidade principal sem antes informar outra como principal!"
        'End If
      End If
    End If
  End If

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").Value >CurrentQuery.FieldByName("DATAFINAL").Value Then
      bsShowMessage("A Data Inicial não pode ser maior que a Data Final", "E")
      CanContinue = False
    End If
  End If


  If(CurrentQuery.FieldByName("OPERACAO").Value = 1)Or(CurrentQuery.FieldByName("OPERACAO").Value = 3)Or(CurrentQuery.FieldByName("OPERACAO").Value = 4)Then
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR_ESPECIALIDADE A WHERE A.PRESTADOR = :PREST AND A.ESPECIALIDADE = :ESPEC AND DATAFINAL IS NULL")
  SQL.ParamByName("PREST").Value = CurrentQuery.FieldByName("PRESTADOR").Value
  SQL.ParamByName("ESPEC").Value = CurrentQuery.FieldByName("ESPECIALIDADE").Value
  SQL.Active = True
  If SQL.FieldByName("HANDLE").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
      bsShowMessage("Data Inicial obrigatória", "E")
      CanContinue = False
    Else

      If CurrentQuery.FieldByName("OPERACAO").Value = 1 Then
        Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
        Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString
        '  Condicao =Condicao +"AND ESPECIALIDADE =  " +CurrentQuery.FieldByName("ESPECIALIDADE").AsString

        Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ESPECIALIDADE", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "ESPECIALIDADE", Condicao)

        If Linha = "" Then
          CanContinue = True
        Else
          CanContinue = False
          bsShowMessage(Linha, "E")
        End If
        Set Interface = Nothing
      End If

    End If
  End If
End If


SQL.Active = False
SQL.Clear
SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC A WHERE A.PRESTADORPROCESSO = :PRESTADORPROCESSO AND A.HANDLE <> :HANDLE AND A.PRINCIPAL = 'S'")
SQL.RequestLive = True
SQL.ParamByName("PRESTADORPROCESSO").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN")
SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.Active = True
If Not SQL.EOF Then
  If CurrentQuery.FieldByName("PRINCIPAL").AsString = "S" Then
    SQL.Edit
    SQL.FieldByName("PRINCIPAL").Value = "N"
    SQL.Post
  End If
Else
  If CurrentQuery.FieldByName("PRINCIPAL").AsString = "N" Then
    'BY WILSON
    'A especialidade principal deixou de ser exigida sendo assim não é mais preciso marcar uma comp principal
    '*****
    'SQL.Clear
    'SQL.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADE A WHERE A.PRESTADOR = :PRESTADOR AND A.PRINCIPAL = 'S'")
    'SQL.ParamByName("PRESTADOR").Value =CurrentQuery.FieldByName("PRESTADOR").AsInteger
    'SQL.Active=True
    'If SQL.EOF Then
    '  CanContinue=False
    '  MsgBox "Nenhuma especialidade marcada como principal"
    'End If
    'END BY WILSON
    'CanContinue=False
    'MsgBox "Nenhuma especialidade marcada como principal"
  End If
End If
SQL.Active = False
'BY WILSON
'NAO PERMITIR CADASTRAR MAIS DE UMA VEZ EM UM MESMO PROCESSO E MESMA OPERACAO  A MESMA ESPECIALIDADE
SQL.Clear
SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_ESPEC A WHERE A.PRESTADORPROCESSO = :PRESTADORPROCESSO")
SQL.Add("AND A.HANDLE <> :HANDLE AND A.OPERACAO = :OPERACAO AND A.ESPECIALIDADE = :ESPECIALIDADE")
SQL.ParamByName("PRESTADORPROCESSO").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN")
SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("OPERACAO").Value = CurrentQuery.FieldByName("OPERACAO").AsString
SQL.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
SQL.Active = True
If Not SQL.EOF Then
  CanContinue = False
  bsShowMessage("Especialidade já cadastrada neste processo!", "E")
  Exit Sub
End If
SQL.Active = False
'END BY WILSON
Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  If Not Ok Then
    bsShowMessage(Mensagem, "E")
    CanContinue = False
  End If
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser

  Dim buscaPrestador As Object
  Set buscaPrestador = NewQuery


  buscaPrestador.Add("SELECT PRESTADOR           ")
  buscaPrestador.Add("  FROM SAM_PRESTADOR_PROC  ")
  buscaPrestador.Add(" WHERE HANDLE = :PROCESSO  ")

  buscaPrestador.ParamByName("PROCESSO").AsInteger = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  buscaPrestador.Active = True

  CurrentQuery.FieldByName("PRESTADOR").AsInteger = buscaPrestador.FieldByName("PRESTADOR").AsInteger

  Set buscaPrestador = Nothing
End Sub

Public Sub BOTAOSELECIONARFILIADO_OnClick()
  Dim Interface As Object
  Dim SQL As Object

  If bsShowMessage("Após selecionar os prestadores filiados, não será possível" + (Chr(13)) + _
             "cadastrar redes restritas neste processo." + (Chr(13)) + _
             "Deseja Continuar ??? ", "Q") = vbYes Then
    Set SQL = NewQuery
    SQL.Add("SELECT P.PRESTADOR FROM SAM_PRESTADOR_PROC P JOIN SAM_PRESTADOR_PROC_CREDEN PC ON (PC.PRESTADORPROCESSO = P.HANDLE) WHERE PC.HANDLE=:HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
    '-------------------------------------------------------
    
    '-------------------------------------------------------
    SQL.Active = False
    SQL.Active = True
    If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
      If CurrentQuery.FieldByName("OPERACAO").Value = 1 Then
        Set Interface = CreateBennerObject("SamProcPrestador.ProcessoPrestador")
        Interface.SelecionaFiliados(CurrentSystem, SQL.FieldByName("PRESTADOR").AsInteger, CurrentQuery.FieldByName("HANDLE").Value)
      Else
        bsShowMessage("Só é permitido selecionar os prestadores filiados" + Chr(10) + _
          "quando a operação for igual a '1-Incluir Especialidade'", "I")
      End If
    Else
      bsShowMessage("Para selecionar os filiados é preciso salvar o registro !!!", "I")
    End If

    Set Interface = Nothing
    Set SQL = Nothing
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOSELECIONARFILIADO" Then
		BOTAOSELECIONARFILIADO_OnClick
	End If
End Sub
