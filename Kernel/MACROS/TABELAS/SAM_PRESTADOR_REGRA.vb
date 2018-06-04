'HASH: 3154735160969196E9AADBBB9B1444A7
'Macro: SAM_PRESTADOR_REGRA
'02/01/2001 - Alterado por Paulo Garcia Junior - liberacao para edição Do registro atraves dos parametros gerais de PRESTADOR
'#Uses "*liberaRegraExcecao"
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
'#Uses "*RegistrarLogAlteracao"

Dim vgEvento As Long


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim Interface  As Object
  Dim vsData     As String
  Dim vsTabela   As String
  Dim vsCampos   As String
  Dim vsColunas  As String
  Dim vsCriterio As String
  Dim viHandle   As Long

  Set Interface = CreateBennerObject("Procura.Procurar")

  vsData    = SQLDate(ServerDate)
  vsTabela  = "SAM_TGE|*SAM_CBHPM[SAM_CBHPM.HANDLE = SAM_TGE.CBHPMTABELA]"
  vsCampos  = "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM"
  vsColunas = "SAM_TGE.ESTRUTURA|SAM_CBHPM.ESTRUTURA|SAM_TGE.DESCRICAO|SAM_CBHPM.DESCRICAO"

  If (NodeInternalCode = 3142) Then 'REGRA
    vsCriterio = "SAM_TGE.HANDLE NOT IN (SELECT GE.EVENTO "
  Else 'EXCEÇÃO
    vsCriterio = "SAM_TGE.HANDLE IN (SELECT GE.EVENTO "
  End If
  vsCriterio = vsCriterio + "                         FROM SAM_ESPECIALIDADEGRUPO_EXEC         GE "
  vsCriterio = vsCriterio + "                         JOIN SAM_ESPECIALIDADEGRUPO              EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO) "
  vsCriterio = vsCriterio + "                         JOIN SAM_ESPECIALIDADE                   E  ON (E.HANDLE = EG.ESPECIALIDADE) "
  vsCriterio = vsCriterio + "                         JOIN SAM_PRESTADOR_ESPECIALIDADE         PE ON (PE.ESPECIALIDADE = E.HANDLE) "
  vsCriterio = vsCriterio + "                         LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.PRESTADORESPECIALIDADE = PE.HANDLE) "
  vsCriterio = vsCriterio + "                        WHERE PE.PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
  vsCriterio = vsCriterio + "                          AND PE.DATAINICIAL <= " + vsData
  vsCriterio = vsCriterio + "                          AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= " + vsData + ") "
  vsCriterio = vsCriterio + "                          AND (PG.ESPECIALIDADEGRUPO = EG.HANDLE OR PG.ESPECIALIDADEGRUPO IS NULL) "
  vsCriterio = vsCriterio + "                          AND GE.EVENTO NOT IN (SELECT X.EVENTO "
  vsCriterio = vsCriterio + "                                                  FROM SAM_PRESTADOR_REGRA X "
  vsCriterio = vsCriterio + "                                                 WHERE X.REGRAEXCECAO   = 'E' "
  vsCriterio = vsCriterio + "                                                   AND X.PERMITERECEBER = 'S' "
  vsCriterio = vsCriterio + "                                                   AND X.PRESTADOR      = PE.PRESTADOR "
  vsCriterio = vsCriterio + "                                                   AND X.DATAINICIAL   <= " + vsData
  vsCriterio = vsCriterio + "                                                   AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vsData + "))) "
  vsCriterio = vsCriterio + "

  viHandle = Interface.Exec(CurrentSystem, vsTabela, vsColunas, 1, vsCampos, vsCriterio, "Eventos que o prestador pode executar", False, EVENTO.Text)
  If (viHandle <> 0) Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").AsInteger = viHandle
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterPost()
  RegistrarLogAlteracao "SAM_PRESTADOR_REGRA", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"
End Sub

Public Sub TABLE_AfterScroll()
  If liberaRegraExcecao <> "" Then
    EVENTO.ReadOnly = True
    EVENTOESTRUTURA.ReadOnly = True
    PERMITEEXECUTAR.ReadOnly = True
    PERMITERECEBER.ReadOnly = True
    PRESTADOR.ReadOnly = True
    REGRAEXCECAO.ReadOnly = True
    TEMPORARIO.ReadOnly = True
  Else
    EVENTO.ReadOnly = False
    EVENTOESTRUTURA.ReadOnly = False
    PERMITEEXECUTAR.ReadOnly = False
    PERMITERECEBER.ReadOnly = False
    PRESTADOR.ReadOnly = False
    REGRAEXCECAO.ReadOnly = False
    TEMPORARIO.ReadOnly = False
  End If

  If CurrentQuery.State = 1 Then
    If Not CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
      DATAINICIAL.ReadOnly = True
    Else
      DATAINICIAL.ReadOnly = False
    End If

    If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
      DATAFINAL.ReadOnly = True
    Else
      DATAFINAL.ReadOnly = False
    End If
  ElseIf CurrentQuery.State = 3 Then
	DATAINICIAL.ReadOnly = False
	DATAFINAL.ReadOnly = False
  End If

  EVENTOESTRUTURA.ReadOnly = True

  If WebMode Then
	If (WebVisionCode = "V_SAM_PRESTADOR_REGRA_668") Then 'REGRA
	  vsCriterio = "A.HANDLE NOT IN (SELECT GE.EVENTO "
	Else 'EXCEÇÃO
	  vsCriterio = "A.HANDLE IN (SELECT GE.EVENTO "
	End If

	vsCriterio = vsCriterio + "                         FROM SAM_ESPECIALIDADEGRUPO_EXEC         GE "
	vsCriterio = vsCriterio + "                         JOIN SAM_ESPECIALIDADEGRUPO              EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO) "
	vsCriterio = vsCriterio + "                         JOIN SAM_ESPECIALIDADE                   E  ON (E.HANDLE = EG.ESPECIALIDADE) "
	vsCriterio = vsCriterio + "                         JOIN SAM_PRESTADOR_ESPECIALIDADE         PE ON (PE.ESPECIALIDADE = E.HANDLE) "
	vsCriterio = vsCriterio + "                         LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.PRESTADORESPECIALIDADE = PE.HANDLE) "
	vsCriterio = vsCriterio + "                        WHERE PE.PRESTADOR = @CAMPO(PRESTADOR)"
	vsCriterio = vsCriterio + "                          AND PE.DATAINICIAL <= " + SQLDate(ServerDate)
	vsCriterio = vsCriterio + "                          AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= " + SQLDate(ServerDate) + ") "
	vsCriterio = vsCriterio + "                          AND (PG.ESPECIALIDADEGRUPO = EG.HANDLE OR PG.ESPECIALIDADEGRUPO IS NULL) "
	vsCriterio = vsCriterio + "                          AND GE.EVENTO NOT IN (SELECT X.EVENTO "
	vsCriterio = vsCriterio + "                                                  FROM SAM_PRESTADOR_REGRA X "
	vsCriterio = vsCriterio + "                                                 WHERE X.REGRAEXCECAO   = 'E' "
	vsCriterio = vsCriterio + "                                                   AND X.PERMITERECEBER = 'S' "
	vsCriterio = vsCriterio + "                                                   AND X.PRESTADOR      = PE.PRESTADOR "
	vsCriterio = vsCriterio + "                                                   AND X.DATAINICIAL   <= " + SQLDate(ServerDate)
	vsCriterio = vsCriterio + "                                                   AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + SQLDate(ServerDate) + "))) "

	EVENTO.WebLocalWhere = vsCriterio
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial ("E", "P", Msg) = "N" Then

    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaRegraExcecao
  If Msg<>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If
  If CanContinue Then
    RegistrarLogAlteracao "SAM_PRESTADOR", CurrentQuery.FieldByName("PRESTADOR").AsInteger, "SAM_PRESTADOR_REGRA.TABLE_BeforeDelete"
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial ("A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaRegraExcecao
  If Msg<>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If

  vgEvento = CurrentQuery.FieldByName("EVENTO").AsInteger

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial ("I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  Msg = liberaRegraExcecao
  If Msg<>"" Then
    CanContinue = False
    bsShowMessage(Msg, "E")
  End If


End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim Interface As Object
  Dim vMsg As String
  Dim vFechar As Boolean

  '---------------------------------------------
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString
  '  Condicao = Condicao + "AND ESPECIALIDADE =  " + CurrentQuery.FieldByName("ESPECIALIDADE").AsString

  If VisibleMode Then
    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_REGRA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", Condicao)

    If Linha = "" Then
      CanContinue = True
    Else
      CanContinue = False
      bsShowMessage(linha + Chr(10) + "Observações: " + Chr(10) + _
                         "- Um evento não pode ser regra e exceção ao mesmo tempo se suas vigências coincidirem;" + Chr(10) + _
                         "- Não é permitido dois registros do mesmo evento com vigências intercaladas.", "E")
    End If
  End If
  Set Interface = Nothing
  '---------------------------------------------

  If VisibleMode Then
	If NodeInternalCode = 3141 Then
	  CurrentQuery.FieldByName("REGRAEXCECAO").Value = "E"
	End If

	If NodeInternalCode = 3142 Then
	  CurrentQuery.FieldByName("REGRAEXCECAO").Value = "R"
	End If
  Else
	If WebMode Then
		If WebVisionCode = "V_SAM_PRESTADOR_REGRA_668" Then
	    	CurrentQuery.FieldByName("REGRAEXCECAO").Value = "R"
	  	Else
	  		If WebVisionCode = "V_SAM_PRESTADOR_REGRA_595" Then
	    		CurrentQuery.FieldByName("REGRAEXCECAO").Value = "E"
	  		End If
	  	End If
	Else
	  If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "" Then
        bsShowMessage("Necessário informar se é ou Regra/Exceção", "E")
	  End If
	End If
  End If

  Dim MensagemErro As String
  If ((CurrentQuery.FieldByName("REGRAEXCECAO").Value = "R") And _
       (CurrentQuery.FieldByName("PERMITEEXECUTAR").AsString <> "S") And _
       (CurrentQuery.FieldByName("PERMITERECEBER").AsString <> "S")) _
       Or _
       ((CurrentQuery.FieldByName("REGRAEXCECAO").Value = "E") And _
       (CurrentQuery.FieldByName("PERMITEEXECUTAR").AsString <> "S") And _
       (CurrentQuery.FieldByName("PERMITERECEBER").AsString <> "S") And _
       (CurrentQuery.FieldByName("PERMITEVISUALIZARCENTRAL").AsString <> "S")) Then
    CanContinue = False
    MensagemErro = "Deve selecionar Permite Executar e/ou Permite Receber"
    If (CurrentQuery.FieldByName("REGRAEXCECAO").Value = "E") Then
      MensagemErro = MensagemErro + " e/ou Permite Visualizar na central"
    End If
    bsShowMessage(MensagemErro, "E")
    Exit Sub
'Adicionado por SMS 77609 - Rodrigo Andrade
  ElseIf ((CurrentQuery.FieldByName("REGRAEXCECAO").Value = "E") And _
          (CurrentQuery.FieldByName("PERMITEEXECUTAR").AsString <> "N") And _
          (CurrentQuery.FieldByName("PERMITEVISUALIZARCENTRAL").AsString <> "S")) Then
	     CanContinue = False
		 MensagemErro = "Não é possivel permitir 'Visualizar na Consulta Prestador x Evento (Central)' se possui uma Exceção de Execução"

		 If WebMode Then
		   MensagemErro = MensagemErro + Chr(13) + "Marque a opção 'Visualizar na Consulta Prestador x Evento (Central)' antes de salvar"
		 End If

		 CurrentQuery.FieldByName("PERMITEVISUALIZARCENTRAL").AsString = "S"
		 bsShowMessage(MensagemErro, "E")
		 Exit Sub
  ElseIf ((CurrentQuery.FieldByName("REGRAEXCECAO").Value = "R") And _
          (CurrentQuery.FieldByName("PERMITEEXECUTAR").AsString <> "S") And _
          (CurrentQuery.FieldByName("PERMITEVISUALIZARCENTRAL").AsString <> "N")) Then
	     CanContinue = False
		 MensagemErro = "Não é possivel permitir 'Visualizar na Consulta Prestador x Evento (Central)' se não possui Regra de Execução"

		 If WebMode Then
		   MensagemErro = MensagemErro + Chr(13) + "Desmarque a opção 'Visualizar na Consulta Prestador x Evento (Central)' antes de salvar"
		 End If

		 CurrentQuery.FieldByName("PERMITEVISUALIZARCENTRAL").AsString = "N"
		 bsShowMessage(MensagemErro, "E")
		 Exit Sub
  End If
  'Fim SMS 77609

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT * FROM SAM_PRESTADOR_REGRAREGIME WHERE REGRA = :REGRA")
  SQL.ParamByName("REGRA").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.Active = True

  If (Not SQL.EOF) And vgEvento <> CurrentQuery.FieldByName("EVENTO").AsInteger Then
    bsShowMessage("Operação inválida !  Esta regra possui regimes de atendimento.", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  Else
    SQL.Clear
    SQL.Add("SELECT * FROM SAM_PRESTADOR_REGRAREDE WHERE REGRA = :REGRA")
    SQL.ParamByName("REGRA").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.Active = True

    If (Not SQL.EOF) And vgEvento <> CurrentQuery.FieldByName("EVENTO").AsInteger Then
      bsShowMessage("Operação inválida !  Esta regra possui redes restritas.", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If


  'Eduardo - 27/12/2004 - SMS 37197
  'Verifica se é necessário fechar As vigências das tabelas de preços
  vFechar = False
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT DATAFINAL          ")
  SQL.Add("  FROM SAM_PRESTADOR_REGRA")
  SQL.Add(" WHERE HANDLE = :HANDLE   ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And (SQL.FieldByName("DATAFINAL").IsNull) Then
    vMsg = "Existe preço relacionado com esta regra nas seguintes tabelas:"
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT COUNT(GE.EVENTO) QTDE                                                        ")
    SQL.Add("  FROM SAM_ESPECIALIDADEGRUPO_EXEC         GE                                       ")
    SQL.Add("  JOIN SAM_ESPECIALIDADEGRUPO              EG ON (EG.HANDLE = GE.ESPECIALIDADEGRUPO)")
    SQL.Add("  JOIN SAM_ESPECIALIDADE                    E ON (E.HANDLE = EG.ESPECIALIDADE)      ")
    SQL.Add("  JOIN SAM_PRESTADOR_ESPECIALIDADE         PE ON (PE.ESPECIALIDADE = E.HANDLE)      ")
    SQL.Add("  LEFT JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.ESPECIALIDADEGRUPO = PE.HANDLE)")
    SQL.Add(" WHERE PE.DATAINICIAL <= :DATA                                                      ")
    SQL.Add("   AND (PE.DATAFINAL IS NULL OR PE.DATAFINAL >= :DATA)                              ")
    SQL.Add("   AND PE.PRESTADOR = :PREsTADOR                                                    ")
    SQL.Add("   AND GE.EVENTO = :EVENTO                                                          ")
    SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
    SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
    SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    '------------------------------------------

    '------------------------------------------
    SQL.Active = True

    If SQL.FieldByName("QTDE").AsInteger = 0 Then
      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT COUNT(D.EVENTO) QTDE                         ")
      SQL.Add("  FROM SAM_PRECOPRESTADOR_DOTAC D                   ")
      SQL.Add(" WHERE D.PRESTADOR = :PRESTADOR                     ")
      SQL.Add("   AND D.DATAINICIAL <= :DATA                       ")
      SQL.Add("   AND (D.DATAFINAL IS NULL OR D.DATAFINAL >= :DATA)")
      SQL.Add("   AND D.EVENTO = :EVENTO                           ")
      SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
      SQL.Active = True
      If SQL.FieldByName("QTDE").AsInteger > 0 Then
        vFechar = True
        vMsg = vMsg + Chr(13) + Chr(10) + "- Tabela de dotações;"
      End If

      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT COUNT(F.EVENTOINICIAL) QTDE                           ")
      SQL.Add("  FROM SAM_PRECOPRESTADOR_FX F                               ")
      SQL.Add(" WHERE F.PRESTADOR = :PRESTADOR                              ")
      SQL.Add("   AND F.DATAINICIAL <= :DATA                                ")
      SQL.Add("   AND (F.DATAFINAL IS NULL OR F.DATAFINAL >= :DATA)         ")
      SQL.Add("   AND (F.EVENTOINICIAL = :EVENTO OR F.EVENTOFINAL = :EVENTO)")
      SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
      SQL.Active = True
      If SQL.FieldByName("QTDE").AsInteger > 0 Then
        vFechar = True
        vMsg = vMsg + Chr(13) + Chr(10) + "- Faixa de eventos;"
      End If

      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT COUNT(D.EVENTO) QTDE                         ")
      SQL.Add("  FROM SAM_PRECOPRESTADORREGIME_DOTAC D             ")
      SQL.Add(" WHERE D.PRESTADOR = :PRESTADOR                     ")
      SQL.Add("   AND D.DATAINICIAL <= :DATA                       ")
      SQL.Add("   AND (D.DATAFINAL IS NULL OR D.DATAFINAL >= :DATA)")
      SQL.Add("   AND D.EVENTO = :EVENTO                           ")
      SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
      SQL.Active = True
      If SQL.FieldByName("QTDE").AsInteger > 0 Then
        vFechar = True
        vMsg = vMsg + Chr(13) + Chr(10) + "- Tabela de dotações no regime de atendimento;"
      End If

      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT COUNT(F.EVENTOINICIAL) QTDE                           ")
      SQL.Add("  FROM SAM_PRECOPRESTADORREGIME_FX F                         ")
      SQL.Add(" WHERE F.PRESTADOR = :PRESTADOR                              ")
      SQL.Add("   AND F.DATAINICIAL <= :DATA                                ")
      SQL.Add("   AND (F.DATAFINAL IS NULL OR F.DATAFINAL >= :DATA)         ")
      SQL.Add("   AND (F.EVENTOINICIAL = :EVENTO OR F.EVENTOFINAL = :EVENTO)")
      SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
      SQL.Active = True
      If SQL.FieldByName("QTDE").AsInteger > 0 Then
        vFechar = True
        vMsg = vMsg + Chr(13) + Chr(10) + "- Faixa de eventos no regime de atendimento;"
      End If

      vMsg = vMsg + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "As vigências nas tabelas relacionadas também serão fechadas."
      vMsg = vMsg + Chr(13) + Chr(10) + "Deseja realmente fechar a vigência?"


      If vFechar Then
        If VisibleMode Then
          If bsShowMessage(vMsg, "Q") = vbYes Then
            vFechar = True
          Else
            vFechar = False
          End If
        Else
          vFechar = True
        End If
      End If


      If vFechar Then
        SQL.Active = False
        SQL.Clear
        SQL.Add("UPDATE SAM_PRECOPRESTADOR_DOTAC                 ")
        SQL.Add("   SET DATAFINAL = :DATA                        ")
        SQL.Add(" WHERE PRESTADOR = :PRESTADOR                   ")
        SQL.Add("   AND DATAINICIAL <= :DATA                     ")
        SQL.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA)")
        SQL.Add("   AND EVENTO = :EVENTO                         ")
        SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
        SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
        SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
        SQL.ExecSQL

        SQL.Clear
        SQL.Add("UPDATE SAM_PRECOPRESTADOR_FX                             ")
        SQL.Add("   SET DATAFINAL = :DATA                                 ")
        SQL.Add(" WHERE PRESTADOR = :PRESTADOR                            ")
        SQL.Add("   AND DATAINICIAL <= :DATA                              ")
        SQL.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA)         ")
        SQL.Add("   AND (EVENTOINICIAL = :EVENTO OR EVENTOFINAL = :EVENTO)")
        SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
        SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
        SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
        SQL.ExecSQL

        SQL.Clear
        SQL.Add("UPDATE SAM_PRECOPRESTADORREGIME_DOTAC           ")
        SQL.Add("   SET DATAFINAL = :DATA                        ")
        SQL.Add(" WHERE PRESTADOR = :PRESTADOR                   ")
        SQL.Add("   AND DATAINICIAL <= :DATA                     ")
        SQL.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA)")
        SQL.Add("   AND EVENTO = :EVENTO                         ")
        SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
        SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
        SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
        SQL.ExecSQL

        SQL.Clear
        SQL.Add("UPDATE SAM_PRECOPRESTADORREGIME_FX                       ")
        SQL.Add("   SET DATAFINAL = :DATA                                 ")
        SQL.Add(" WHERE PRESTADOR = :PRESTADOR                            ")
        SQL.Add("   AND DATAINICIAL <= :DATA                              ")
        SQL.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA)         ")
        SQL.Add("   AND (EVENTOINICIAL = :EVENTO OR EVENTOFINAL = :EVENTO)")
        SQL.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
        SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
        SQL.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
        SQL.ExecSQL
        'Else
        '  CanContinue = False

      End If
    End If
  End If
  'fim SMS 37197


  'SQL.Add("SELECT COUNT(*) T")
  'SQL.Add("  FROM SAM_PRESTADOR_REGRA X")
  'SQL.Add(" WHERE X.PRESTADOR = :P")
  'SQL.Add("   AND X.EVENTO = :E")
  'SQL.Add("   AND X.REGRAEXCECAO IN ('R','E')")

  'SQL.ParamByName("P").Value = RecordHandleOfTable("SAM_PRESTADOR")
  'SQL.ParamByName("E").Value = CurrentQuery.FieldByName("EVENTO").AsInteger

  'SQL.Active = True
  'If SQL.FieldByName("T").AsInteger > 1 Then
  '	CanContinue = False
  '	MsgBox "O evento não pode ser registrado como Regra e exceção ao mesmo tempo"
  'End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------

Public Function checkPermissaoFilial (pServico As String, pTabela As String, pMsg As String) As String

  Dim vFiltro, vResultado As String
  Dim qAuxiliar, qPermissoes, SamPrestadorParametro As Object
  Dim qFilialProc As Object

  Dim SamPrestador

  Set qPermissoes = NewQuery
  Set qFilialProc = NewQuery
  Set qAuxiliar = NewQuery
  Set SamPrestadorParametro = NewQuery

  SamPrestadorParametro.Add (" SELECT CONTROLEDEACESSO, BLOQUEIOFILIALPROCESSAMENTO ")
  SamPrestadorParametro.Add (" FROM SAM_PARAMETROSPRESTADOR  ")
  SamPrestadorParametro.Active = True

  If SamPrestadorParametro.FieldByName("CONTROLEDEACESSO").AsString = "N" Then
    checkPermissaoFilial = "(SELECT HANDLE FROM MUNICIPIOS)"
    pMsg = ""
    Exit Function
  End If

  ' começa o controle de acesso
  ' verifica bloqueio filial de processamento

  If SamPrestadorParametro.FieldByName("BLOQUEIOFILIALPROCESSAMENTO").AsString = "S" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT FILIALPADRAO")
    qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
    qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
    qAuxiliar.Active = True
    qFilialProc.Active = False
    qFilialProc.Add("Select FILIALPROCESSAMENTO FROM FILIAIS WHERE HANDLE = :HANDLE")
    qFilialProc.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("FILIALPADRAO").Value
    qFilialProc.Active = True
    If qFilialProc.FieldByName("FILIALPROCESSAMENTO").AsInteger = qAuxiliar.FieldByName("FILIALPADRAO").Value Then
      checkPermissaoFilial = "N"
      pMsg = "Permissão negada! Filial padrão do usuário igual a sua filial de processamento."
      Exit Function
    End If
  End If

  qAuxiliar.Active = False
  qAuxiliar.Clear
  If pTabela = "P" Then
    ' se For alterar os dados de um PRESTADOR já cadastrado
    If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
      qAuxiliar.Add("SELECT FILIALPADRAO")
      qAuxiliar.Add("  FROM SAM_PRESTADOR")
      qAuxiliar.Add(" WHERE HANDLE = :HPRESTADOR")
      qAuxiliar.ParamByName("HPRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      qAuxiliar.Active = True
      If qAuxiliar.FieldByName("FILIALPADRAO").IsNull Then
        qAuxiliar.Clear
        qAuxiliar.Add("SELECT FILIALPADRAO")
        qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
        qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
        qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
        qAuxiliar.Active = True
      End If
      ' se For cadastrar um novo PRESTADOR
    Else
      qAuxiliar.Add("SELECT FILIALPADRAO")
      qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
      qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
      qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
      qAuxiliar.Active = True
    End If

    qPermissoes.Active = False
    qPermissoes.Clear
    qPermissoes.Add("SELECT X.ALTERAR, ")
    qPermissoes.Add("       X.INCLUIR, ")
    qPermissoes.Add("       X.EXCLUIR, ")
    qPermissoes.Add("       X.FILIAL ")
    qPermissoes.Add("  FROM (SELECT A.ALTERAR ALTERAR, ")
    qPermissoes.Add("               A.INCLUIR INCLUIR, ")
    qPermissoes.Add("               A.EXCLUIR EXCLUIR, ")
    qPermissoes.Add("               A.FILIAL  FILIAL ")
    qPermissoes.Add("          FROM Z_GRUPOUSUARIOS_FILIAIS A ")
    qPermissoes.Add("         WHERE  A.USUARIO = :USUARIO ")
    qPermissoes.Add("           AND  A.FILIAL  = :FILIAL ")
    qPermissoes.Add("        UNION ")
    qPermissoes.Add("        SELECT U.ALTERAR      ALTERAR, ")
    qPermissoes.Add("               U.INCLUIR      INCLUIR, ")
    qPermissoes.Add("               U.EXCLUIR      EXCLUIR, ")
    qPermissoes.Add("               U.FILIALPADRAO FILIAL ")
    qPermissoes.Add("          FROM Z_GRUPOUSUARIOS U ")
    qPermissoes.Add("         WHERE U.HANDLE = :USUARIO ")
    qPermissoes.Add("           AND U.FILIALPADRAO  = :FILIAL) X ")

    qPermissoes.ParamByName("USUARIO").Value = CurrentUser
    qPermissoes.ParamByName("FILIAL").Value = qAuxiliar.FieldByName("FILIALPADRAO").AsInteger
    qPermissoes.Active = True
  End If


  If pServico = "A" Then
    ' Verifica se pode alterar conforme a filial padrao
    vFiltro = vFiltro + _
              "SELECT DISTINCT M.HANDLE " + _
              "  FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "       SAM_REGIAO R, " + _
              "       MUNICIPIOS M " + _
              " WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "   AND R.FILIAL = A.FILIAL " + _
              "   AND M.REGIAO = R.HANDLE " + _
              "   AND A.ALTERAR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.ALTERAR = 'S'  "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode alterar
    vFiltro = ""
    vFiltro = vFiltro + _
              "SELECT DISTINCT M.HANDLE " + _
              "   FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "        MUNICIPIOS M, " + _
              "        SAM_REGIAO R " + _
              "  WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "    AND M.REGIAO = R.HANDLE " + _
              "    AND A.FILIAL = R.FILIAL " + _
              "    AND A.ALTERAR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.ALTERAR = 'S' "
  End If
  If pServico = "I" Then
    ' Verifica se pode incluir conforme a filial padrao
    vFiltro = vFiltro + _
              "SELECT DISTINCT M.HANDLE " + _
              "  FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "       SAM_REGIAO R, " + _
              "       MUNICIPIOS M " + _
              " WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "   AND R.FILIAL = A.FILIAL " + _
              "   AND M.REGIAO = R.HANDLE " + _
              "   AND A.INCLUIR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.INCLUIR = 'S'  "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode incluir
    vFiltro = ""
    vFiltro = vFiltro + _
              "Select DISTINCT M.HANDLE " + _
              "   FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "        MUNICIPIOS M, " + _
              "        SAM_REGIAO R " + _
              "  WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "    AND M.REGIAO = R.HANDLE " + _
              "    AND A.FILIAL = R.FILIAL " + _
              "    AND A.INCLUIR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE " + _
              "    FROM Z_GRUPOUSUARIOS U, " + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.INCLUIR = 'S'  "

  End If

  ' se não estiver cadastrado
  If (qPermissoes.FieldByName("ALTERAR").IsNull) Then
    If pServico = "" Then
      checkPermissaoFilial = ""
      Exit Function
    End If
  End If

  ' se não informou o servico, retorna uma String com os servicos permitidos "LAIE"
  If (pServico = "") Then
    vResultado = ""
    If qPermissoes.FieldByName("ALTERAR").AsString = "S" Then
      vResultado = vResultado + "A"
    End If
    If qPermissoes.FieldByName("INCLUIR").AsString = "S" Then
      vResultado = vResultado + "I"
    End If
    If qPermissoes.FieldByName("EXCLUIR").AsString = "S" Then
      vResultado = vResultado + "E"
    End If
    ' se informou o servico, retorna S/N
  Else
    Select Case pServico
      Case "A"
        If qPermissoes.FieldByName("ALTERAR").AsString = "S" Then
          vResultado = "S"
          If (Not qAuxiliar.FieldByName("Handle").IsNull) Then
            vResultado = vFiltro
          Else
            vResultado = "N"
            pMsg = "Permissão negada! Usuário não pode alterar."
          End If
        Else
          vResultado = "N"
          pMsg = "Permissão negada! Usuário não pode alterar."
        End If
      Case "I"
        If qPermissoes.FieldByName("INCLUIR").AsString = "S" Then
          vResultado = "S"
          If (Not qAuxiliar.FieldByName("Handle").IsNull) Then
            vResultado = vFiltro
          Else
            vResultado = "N"
            pMsg = "Permissão negada! Usuário não pode incluir."
          End If
        Else
          vResultado = "N"
          pMsg = "Permissão negada! Usuário não pode incluir."
        End If
      Case "E"
        If qPermissoes.FieldByName("EXCLUIR").AsString = "S" Then
          vResultado = "S"
        Else
          vResultado = "N"
          pMsg = "Permissão negada! Usuário não pode excluir."
        End If
    End Select
  End If
  checkPermissaoFilial = vResultado
End Function


Public Function BuscarFiliais(prFilial As Long, prFilialProcessamento As Long, prMsg As String) As Boolean

  Dim qPermissoes As Object
  Set qPermissoes = NewQuery

  BuscarFiliais = True
  qPermissoes.Active = False
  qPermissoes.Clear
  qPermissoes.Add("SELECT A.HANDLE, A.FILIALPROCESSAMENTO")
  qPermissoes.Add("FROM   Z_GRUPOUSUARIOS U,             ")
  qPermissoes.Add("       FILIAIS A                      ")
  qPermissoes.Add("WHERE  (U.HANDLE = :USUARIO)          ")
  qPermissoes.Add("AND    (A.HANDLE = U.FILIALPADRAO)    ")
  qPermissoes.ParamByName("USUARIO").Value = CurrentUser
  qPermissoes.Active = True

  If qPermissoes.EOF Then
    prMsg = "Problemas Usuario x Filial."
    Exit Function
  End If

  prFilial = qPermissoes.FieldByName("HANDLE").AsInteger
  prFilialProcessamento = qPermissoes.FieldByName("FILIALPROCESSAMENTO").AsInteger
  prMsg = ""
  BuscarFiliais = False
  Set qPermissoes = Nothing
End Function
