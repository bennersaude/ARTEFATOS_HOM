'HASH: 97A890C81ED7156A980D25E916FF6E81
'Macro: SAM_PRESTADOR_PROC_REGEXC

'#Uses "*bsShowMessage"

'Mauricio Ibelli -04/01/2002 -sms3165 -Se filial padrao do prestador for nulo não checar responsavel'
'Claudemir -05/08/2002 -sms6362 -Vigências nas regras e exceções
Dim Mensagem As String

Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery

  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True


  SQL.Add("SELECT SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.FILIALPADRAO FROM SAM_PRESTADOR_PROC, SAM_PRESTADOR WHERE SAM_PRESTADOR_PROC.HANDLE = :HANDLE And  SAM_PRESTADOR.HANDLE = SAM_PRESTADOR_PROC.PRESTADOR")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")
  SQL.Active = True
  Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And((SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser)Or(SQL.FieldByName("FILIALPADRAO").IsNull)), True, False)

  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida" + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <>CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub EVENTO_OnExit()
  If CurrentQuery.FieldByName("OPERACAO").IsNull And Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    bsShowMessage("Selecione uma operação.", "E")
    CurrentQuery.FieldByName("EVENTO").Clear
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT ULTIMONIVEL FROM SAM_TGE WHERE HANDLE = " + CurrentQuery.FieldByName("EVENTO").AsString)
    SQL.Active = True

    If SQL.FieldByName("ULTIMONIVEL").AsString <>"S" Then
      bsShowMessage("O evento deve ser último nível.", "E")
      CurrentQuery.FieldByName("EVENTO").Clear
      EVENTO.SetFocus
    End If

    If CurrentQuery.FieldByName("OPERACAO").AsString <> "I" Then

      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR_REGRA WHERE EVENTO = :HANDLE")
      SQL.Add("  AND PRESTADOR = :PRESTADOR                                 ")

      If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "R" Then
        SQL.Add("AND REGRAEXCECAO = 'R'")
      Else
        SQL.Add("AND REGRAEXCECAO = 'E'")
      End If
      SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ParamByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
      SQL.Active = True

      If SQL.FieldByName("HANDLE").IsNull Then
        If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "R" Then
          bsShowMessage("Este evento não pertence as regras do prestador !", "I")
        Else
          bsShowMessage("Este evento não pertence as exceções do prestador !", "I")
        End If
        CurrentQuery.FieldByName("EVENTO").Clear
        EVENTO.SetFocus
        Exit Sub
      End If
    End If

    Set SQL = Nothing

    If CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then
      DATAINICIAL1.ReadOnly = False
      DATAFINAL1.ReadOnly = False
    End If

  End If
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  Dim Interface, SQL As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False

  If CurrentQuery.FieldByName("OPERACAO").IsNull Then
    MsgBox "Selecione uma operação!"
    Exit Sub
  End If

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_TGE.NIVELAUTORIZACAO"
  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
    vHandle = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTO.LocateText)
  Else

    If CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then
      DATAINICIAL1.ReadOnly = False
      DATAFINAL1.ReadOnly = False
    End If

    Set SQL = NewQuery
    SQL.Add("SELECT * FROM SAM_PRESTADOR WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR")
    SQL.Active = True
    If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "R" Then
      vTitulo = "Eventos com Regras no Prestador " + SQL.FieldByName("NOME").AsString
    Else
      vTitulo = "Eventos com Exceções no Prestador " + SQL.FieldByName("NOME").AsString
    End If
    vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_TGE.NIVELAUTORIZACAO|SAM_PRESTADOR_REGRA.DATAINICIAL|SAM_PRESTADOR_REGRA.DATAFINAL"
    vCampos = "Evento|Descrição|Nível|Data inicial|Data final"
    vCriterio = "SAM_TGE.ULTIMONIVEL = 'S' "
    vCriterio = vCriterio + " AND SAM_PRESTADOR_REGRA.PRESTADOR = " + SQL.FieldByName("HANDLE").AsString
    vCriterio = vCriterio + " AND SAM_PRESTADOR_REGRA.REGRAEXCECAO = '" + CurrentQuery.FieldByName("REGRAEXCECAO").AsString + "'"

    vPrestadorRegra = Interface.Exec(CurrentSystem, "SAM_PRESTADOR_REGRA|SAM_TGE[SAM_PRESTADOR_REGRA.EVENTO = SAM_TGE.HANDLE]", vColunas, 1, vCampos, vCriterio, vTitulo, True, EVENTO.LocateText)

    If vPrestadorRegra <>0 Then
      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT E.HANDLE FROM SAM_TGE E WHERE E.HANDLE = (SELECT PR.EVENTO FROM SAM_PRESTADOR_REGRA PR WHERE PR.HANDLE = :HANDLE) ")
      SQL.ParamByName("HANDLE").Value = vPrestadorRegra
      SQL.Active = True
      vHandle = SQL.FieldByName("HANDLE").Value

      CurrentQuery.FieldByName("PRESTADORREGRA").Value = vPrestadorRegra

      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT DATAINICIAL, DATAFINAL")
      SQL.Add("  FROM SAM_PRESTADOR_REGRA   ")
      SQL.Add(" WHERE HANDLE = :HANDLE      ")
      SQL.ParamByName("HANDLE").Value = vPrestadorRegra
      SQL.Active = True
      CurrentQuery.FieldByName("DATAINICIAL1").Value = SQL.FieldByName("DATAINICIAL").Value
      CurrentQuery.FieldByName("DATAFINAL1").Value = SQL.FieldByName("DATAFINAL").Value

      REGRAEXCECAO.ReadOnly = True
      OPERACAO.ReadOnly = True

    End If
    Set SQL = Nothing
  End If

  CurrentQuery.Edit
  If vHandle <>0 Then
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  Else
    CurrentQuery.FieldByName("EVENTO").Clear
  End If
  Set Interface = Nothing

  OPERACAO.ReadOnly = True
  If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then 'incluir
    DATAINICIAL.ReadOnly = False
    DATAINICIAL1.ReadOnly = True
    DATAFINAL.ReadOnly = False
    DATAFINAL1.ReadOnly = True
  ElseIf CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then 'alterar
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
  Else 'excluir
    DATAINICIAL.ReadOnly = True
    DATAINICIAL1.ReadOnly = True
    DATAFINAL.ReadOnly = True
    DATAFINAL1.ReadOnly = True
  End If

End Sub


Public Sub TABLE_AfterInsert()
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_REGEXC"
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    REGRAEXCECAO.ReadOnly = True
    OPERACAO.ReadOnly = True
  Else
    REGRAEXCECAO.ReadOnly = False
    OPERACAO.ReadOnly = False
  End If
  If CurrentQuery.FieldByName("OPERACAO").IsNull Then
    DATAINICIAL.ReadOnly = False
    DATAFINAL.ReadOnly = False
  Else
    If CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then
      DATAINICIAL1.ReadOnly = False
      DATAFINAL1.ReadOnly = False
      DATAINICIAL.ReadOnly = True
      DATAFINAL.ReadOnly = True
    Else
      DATAINICIAL1.ReadOnly = True
      DATAFINAL1.ReadOnly = True
      If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
        DATAINICIAL.ReadOnly = False
        DATAFINAL.ReadOnly = False
      End If
      If CurrentQuery.FieldByName("OPERACAO").AsString = "E" Then
        DATAINICIAL.ReadOnly = True
        DATAFINAL.ReadOnly = True
      End If
    End If
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


    If CurrentQuery.FieldByName("OPERACAO").IsNull And Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    bsShowMessage("Selecione uma operação.", "E")
    CurrentQuery.FieldByName("EVENTO").Clear
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT ULTIMONIVEL FROM SAM_TGE WHERE HANDLE = " + CurrentQuery.FieldByName("EVENTO").AsString)
    SQL.Active = True

    If SQL.FieldByName("ULTIMONIVEL").AsString <>"S" Then
      bsShowMessage("O evento deve ser último nível.", "E")
      CurrentQuery.FieldByName("EVENTO").Clear
      EVENTO.SetFocus
    End If

    If CurrentQuery.FieldByName("OPERACAO").AsString <> "I" Then

      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR_REGRA WHERE EVENTO = :HANDLE")
      SQL.Add("  AND PRESTADOR = :PRESTADOR                                 ")

      If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "R" Then
        SQL.Add("AND REGRAEXCECAO = 'R'")
      Else
        SQL.Add("AND REGRAEXCECAO = 'E'")
      End If
      SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
      SQL.ParamByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
      SQL.Active = True

      If SQL.FieldByName("HANDLE").IsNull Then
        If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "R" Then
          bsShowMessage("Este evento não pertence as regras do prestador !", "I")
        Else
          bsShowMessage("Este evento não pertence as exceções do prestador !", "I")
        End If
        CurrentQuery.FieldByName("EVENTO").Clear
        EVENTO.SetFocus
        Exit Sub
      End If
    End If

    Set SQL = Nothing

    If CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then
      DATAINICIAL1.ReadOnly = False
      DATAFINAL1.ReadOnly = False
    End If

  End If


  If CurrentQuery.FieldByName("REGRAEXCECAO").AsString = "R" Then
    CanContinue = VerRegra
  Else
    CanContinue = VerExcecao
  End If

  If CanContinue = True Then

    If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
      '---------------------------------------------
      Dim Interface As Object
      Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
      Condicao = " AND SAM_PRESTADOR_PROC_REGEXC.PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString + " AND SAM_PRESTADOR_PROC_REGEXC.HANDLE IN (SELECT SAM_PRESTADOR_PROC_REGEXC.HANDLE " + _
                 " FROM SAM_PRESTADOR_PROC, " + _
                 " SAM_PRESTADOR_PROC_CREDEN, " + _
                 " SAM_PRESTADOR_PROC_REGEXC " + _
                 " WHERE SAM_PRESTADOR_PROC_REGEXC.PRESTADOR           = " + CurrentQuery.FieldByName("PRESTADOR").AsString + _
                 " And SAM_PRESTADOR_PROC_REGEXC.PRESTADORPROCESSO   = SAM_PRESTADOR_PROC_CREDEN.HANDLE" + _
                 " And SAM_PRESTADOR_PROC_CREDEN.PRESTADORPROCESSO   = SAM_PRESTADOR_PROC.HANDLE" + _
                 " And SAM_PRESTADOR_PROC.PRESTADOR                  = " + CurrentQuery.FieldByName("PRESTADOR").AsString + _
                 " And SAM_PRESTADOR_PROC.DATAFINAL Is Null AND SAM_PRESTADOR_PROC_REGEXC.HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString + ")"
      Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_REGEXC", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", Condicao)
      Set Interface = Nothing
      If linha <>"" Then
        bsShowMessage(linha + Chr(10) + "Observações: " + Chr(10) + _
                           "- Um evento não pode ser regra e exceção ao mesmo tempo se suas vigências coincidirem;" + Chr(10) + _
                           "- Não é permitido dois registros do mesmo evento com vigências intercaladas.", "E")
        CanContinue = False
      End If
    Else
      If CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then
        If Not CurrentQuery.FieldByName("DATAINICIAL1").IsNull Then
          Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
          Condicao = " AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString
          Condicao = Condicao + " AND HANDLE <> " + CurrentQuery.FieldByName("PRESTADORREGRA").AsString
          Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_REGRA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL1").AsDateTime, CurrentQuery.FieldByName("DATAFINAL1").AsDateTime, "EVENTO", Condicao)
          Set Interface = Nothing
          If linha <>"" Then
            bsShowMessage(linha, "E")
            CanContinue = False
          End If
        Else
          bsShowMessage("A data inicial na nova vigência não pode ser nula", "E")
          CanContinue = False
        End If
      End If
    End If

  End If
End Sub

Public Function VerRegra As Boolean
  VerRegra = True
  Dim Interface As Object
  Dim SQL As Object


  '---------------------------------------------
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString
  Condicao = Condicao + " AND REGRAEXCECAO =  'R'"

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_REGRA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", Condicao)

  If Linha <> "" Then
    If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
      VerRegra = False
      bsShowMessage(linha + Chr(10) + " " + Chr(10) + _
                         "Já existe um registro deste evento com vigência aberta nas regras do prestador (Informações cadastrais - Regras)", "E")
    End If
  End If

  If Linha = "" Then
    Condicao = Condicao + " AND PRESTADORPROCESSO = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_REGEXC", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", Condicao)

    If Linha <> "" Then
      If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
        VerRegra = False
        bsShowMessage(linha + Chr(10) + " " + Chr(10) + _
                           "Já existe um registro deste evento com vigência aberta ou que coincida com a informada", "E")
      End If
    End If
  End If

  Set Interface = Nothing

End Function

Public Function VerExcecao As Boolean
  VerExcecao = True
  Dim Interface As Object

  '---------------------------------------------
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Condicao = "AND PRESTADOR     = " + CurrentQuery.FieldByName("PRESTADOR").AsString
  Condicao = Condicao + " AND REGRAEXCECAO =  'E'"

  Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_REGRA", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", Condicao)

  If Linha <> "" Then
    If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
      VerExcecao = False
      bsShowMessage(linha + Chr(10) + " " + Chr(10) + _
                         "Já existe um registro deste evento com vigência aberta nas exceções do prestador (Informações cadastrais - Exceções)", "E")
    End If
  End If

  If Linha <> "" Then
    Condicao = Condicao + " AND PRESTADORPROCESSO = " + CurrentQuery.FieldByName("PRESTADORPROCESSO").AsString
    Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_PROC_REGEXC", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", Condicao)

    If Linha <> "" Then
      If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
        VerExcecao = False
        bsShowMessage(linha + Chr(10) + " " + Chr(10) + _
                           "Já existe um registro deste evento com vigência aberta ou que coincida com a informada", "E")
      End If
    End If
  End If

  Set Interface = Nothing
End Function

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

  CanContinue = Ok
  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
  End If

End Sub


Public Sub TABLE_BeforeScroll()
  Dim SQL, SQL1 As Object

  DATAINICIAL1.ReadOnly = True
  DATAFINAL1.ReadOnly = True

  EVENTO.ReadOnly = False
  DATAINICIAL.ReadOnly = False
  DATAFINAL.ReadOnly = False

  If CurrentQuery.FieldByName("OPERACAO").AsString <>"I" Then
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
  End If

  If CurrentQuery.FieldByName("OPERACAO").AsString = "A" Then
    DATAINICIAL1.ReadOnly = False
    DATAFINAL1.ReadOnly = False
  Else
  End If

  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    Set SQL = NewQuery
    SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_REGEXC_REDE WHERE PRESTADORPROCREGEXC = :PRESTADORPROCREGEXC")
    SQL.ParamByName("PRESTADORPROCREGEXC").Value = CurrentQuery.FieldByName("HANDLE").Value
    SQL.Active = True
    If Not SQL.EOF Then
      EVENTO.ReadOnly = True
      DATAINICIAL1.ReadOnly = True
      DATAFINAL1.ReadOnly = True
      DATAINICIAL.ReadOnly = True
      DATAFINAL.ReadOnly = True
    Else
      EVENTO.ReadOnly = False
      Set SQL1 = NewQuery
      SQL1.Add("SELECT * FROM SAM_PRESTADOR_PROC_REGEXC_REG WHERE PRESTADORPROCREGEXC = :PRESTADORPROCREGEXC")
      SQL1.ParamByName("PRESTADORPROCREGEXC").Value = CurrentQuery.FieldByName("HANDLE").Value
      SQL1.Active = True
      If Not SQL1.EOF Then
        EVENTO.ReadOnly = True
        DATAINICIAL1.ReadOnly = True
        DATAFINAL1.ReadOnly = True
        DATAINICIAL.ReadOnly = True
        DATAFINAL.ReadOnly = True
      End If
    End If
    Set SQL = Nothing
    Set SQL1 = Nothing
  End If

End Sub



Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser

  If WebMode Then
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
  End If
End Sub

Public Sub BOTAOSELECIONARFILIADOS_OnClick()
  Dim Interface As Object
  Dim SQL As Object

  If CurrentQuery.State <>3 Then
    If CurrentQuery.FieldByName("OPERACAO").AsString = "I" Then
      If bsShowMessage("Após selecionar os prestadores filiados, não será possível" + (Chr(13)) + _
                 "cadastrar redes restritas neste processo." + (Chr(13)) + _
                 "Deseja Continua ??? ", "Q") = vbYes Then
        Set SQL = NewQuery
        SQL.Add("SELECT P.PRESTADOR FROM SAM_PRESTADOR_PROC P JOIN SAM_PRESTADOR_PROC_CREDEN PC ON (PC.PRESTADORPROCESSO = P.HANDLE) WHERE PC.HANDLE=:HANDLE")
        SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
        '-------------------------------------------------------

        '-------------------------------------------------------
        SQL.Active = False
        SQL.Active = True

        Set Interface = CreateBennerObject("SamProcPrestador.ProcessoPrestador")
        Interface.SelecionaFiliadosRegraExc(CurrentSystem, SQL.FieldByName("PRESTADOR").AsInteger, RecordHandleOfTable("SAM_PRESTADOR_PROC_REGEXC"))

        Set Interface = Nothing
        Set SQL = Nothing
      End If
    Else
      bsShowMessage("Só é permitido selecionar prestadores filiados para 'Operação' do tipo 'Incluir' !", "E")
    End If
  Else
    bsShowMessage("Só é permitido selecionar prestadores filiados após confirmar o cadastro !", "E")
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOSELECIONARFILIADOS" Then
		BOTAOSELECIONARFILIADOS_OnClick
	End If
End Sub
