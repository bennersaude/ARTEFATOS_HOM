'HASH: B6550FC1393423B233E4B50412EEEB43
'SAM_PROCREF_PRESTADOR

Option Explicit

'#Uses "*ProcuraPrestador"
'#Uses "*bsShowMessage"

Dim vBotaoAprovar As String
Dim qParametro As Object

Public Sub BOTAOAPROVAR_OnClick()
  Dim q1 As Object

  If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then
    bsShowMessage("Registro em inserção ou edição !", "I")
    Exit Sub
  End If


  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PROCREF WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If q1.FieldByName("SITUACAO").AsString = "F" Then
    bsShowMessage("Este processo de avaliação está fechado!", "I")
    Exit Sub
  ElseIf q1.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Este processo de avaliação está cancelado!", "I")
    Exit Sub
  End If


  q1.Active = False
  q1.Clear
  q1.Add("SELECT REVERTER FROM SAM_MOTIVOREFERENCIAMENTO WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOREFERENCIAMENTO").Value
  q1.Active = True

  If q1.FieldByName("REVERTER").AsString = "N" Then
    bsShowMessage("O motivo de referenciamento escolhido não permite reverter situação!", "I")
    Exit Sub
  End If

  CurrentQuery.Edit
  vBotaoAprovar = "S"
  CurrentQuery.FieldByName("SITUACAO").AsString = "A"
  CurrentQuery.FieldByName("APROVACAODATA").Value = ServerNow
  CurrentQuery.FieldByName("APROVACAOUSUARIO").Value = CurrentUser
  If Not CurrentQuery.FieldByName("OBSERVACAO").IsNull Then
    CurrentQuery.FieldByName("OBSERVACAO").Value = CurrentQuery.FieldByName("OBSERVACAO").Value + Chr(13) + "Este prestador estava não-referenciável, mas foi revertido."
  Else
    CurrentQuery.FieldByName("OBSERVACAO").Value = "Este prestador estava não-referenciável, mas foi revertido."
  End If
  APROVACAOJUSTIFICATIVA.ReadOnly = False

End Sub

Public Sub BOTAOEXCLUIR_OnClick()
  Dim qExec As Object
  Dim q1 As Object

  On Error GoTo FIM

  If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then
    bsShowMessage("Registro em inserção ou edição !", "I")
    Exit Sub
  End If

  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PROCREF WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If q1.FieldByName("SITUACAO").AsString = "F" Then
    bsShowMessage("Este processo de avaliação está finalizado!", "I")
    Exit Sub
  ElseIf q1.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Este processo de avaliação está cancelado!", "I")
    Exit Sub
  End If


  If Not CurrentQuery.FieldByName("FINALIZACAODATA").IsNull Then
    bsShowMessage("Esta avaliação está finalizada!", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("GERACAODATA").IsNull Then
    bsShowMessage("Os dados das avaliações deste prestador ainda não foram gerados !", "I")
    Exit Sub
  End If
  If bsShowMessage("Deseja realmente excluir os dados das avaliações deste prestador ?", "Q") = vbYes Then

    Set qExec = NewQuery

    If Not InTransaction Then
      StartTransaction
    End If

    qExec.Add("delete sam_procref_prestador_item_res                                                                        ")
    qExec.Add(" where procprestavalitem in (select i.handle                                                                 ")
    qExec.Add("                               from sam_procref_prestador_aval_ite i                                         ")
    qExec.Add("                              where i.processorefprestaval in (select x.handle                               ")
    qExec.Add("                                                                 from sam_procref_prestador_aval x           ")
    qExec.Add("                                                                where x.procrefprestador = :procrefprestador)")
    qExec.Add("                            )                                                                                ")
    qExec.ParamByName("procrefprestador").Value = CurrentQuery.FieldByName("handle").Value
    qExec.ExecSQL

    qExec.Clear
    qExec.Add("delete sam_procref_prestador_aval_ite                                         ")
    qExec.Add(" where processorefprestaval in (select x.handle                               ")
    qExec.Add("                                  from sam_procref_prestador_aval x           ")
    qExec.Add("                                 where x.procrefprestador = :procrefprestador)")
    qExec.ParamByName("procrefprestador").Value = CurrentQuery.FieldByName("handle").Value
    qExec.ExecSQL

    qExec.Clear
    qExec.Add("delete sam_procref_prestador_aval where procrefprestador = :procrefprestador")
    qExec.ParamByName("procrefprestador").Value = CurrentQuery.FieldByName("handle").Value
    qExec.ExecSQL

    qExec.Clear
    qExec.Add("delete sam_procref_prestador_espec where procrefprestador = :procrefprestador")
    qExec.ParamByName("procrefprestador").Value = CurrentQuery.FieldByName("handle").Value
    qExec.ExecSQL

    Set qExec = Nothing

    If InTransaction Then
      Commit
    End If


    CurrentQuery.Edit
    CurrentQuery.FieldByName("GERACAODATA").Value = Null
    CurrentQuery.FieldByName("GERACAOUSUARIO").Value = Null
    CurrentQuery.Post

    bsShowMessage("Processo concluído !", "I")

    RefreshNodesWithTable("SAM_PROCREF_PRESTADOR")

  End If

  Exit Sub
FIM:
  If InTransaction Then
    Rollback
  End If


End Sub

Public Sub BOTAOFINALIZAR_OnClick()
  Dim q1 As Object
  Dim S1 As String

  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PROCREF WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If q1.FieldByName("SITUACAO").AsString = "F" Then
    bsShowMessage("Este processo de avaliação está FECHADO!", "I")
    Exit Sub
  ElseIf q1.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Este processo de avaliação está cancelado!", "I")
    Exit Sub
  End If


  If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then
    bsShowMessage("Registro em inserção ou edição !", "I")
    Exit Sub
  End If


  If VisibleMode Or WebMode Then
    If CurrentQuery.FieldByName("GERACAODATA").IsNull Then
      bsShowMessage("Não existe informação de avaliação para esta avaliação!", "I")
      Exit Sub
    End If

    If Not CurrentQuery.FieldByName("FINALIZACAODATA").IsNull Then
      bsShowMessage("Esta avaliação já se encontra finalizada !", "I")
      Exit Sub
    End If


    q1.Active = False
    q1.Clear
    q1.Add("SELECT DISTINCT A.DESCRICAO                   ")
    q1.Add("  FROM SAM_AVALIACAOREF                 A,    ")
    q1.Add("       SAM_PROCREF_PRESTADOR_AVAL      PA     ")
    q1.Add(" WHERE PA.AVALIACAOREF = A.HANDLE             ")
    q1.Add("   AND PA.PROCREFPRESTADOR = :HANDLE          ")
    q1.Add("   AND PA.SITUACAO = 'E'                      ")
    q1.Add("   AND NOT EXISTS(SELECT X.HANDLE             ")
    q1.Add("                    FROM SAM_PROCREF_PRESTADOR_AVAL X")
    q1.Add("                   WHERE SITUACAO = 'R' AND REPROVACAO = 'N' AND PROCREFPRESTADOR = :HANDLE)")
    q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    q1.Active = True

    If Not q1.EOF Then
      S1 = "As  avaliações: " + Chr(13)
      S1 = S1 + q1.FieldByName("DESCRICAO").AsString
      q1.Next
      While Not q1.EOF
        S1 = S1 + Chr(13) + q1.FieldByName("DESCRICAO").AsString
        q1.Next
      Wend
      S1 = S1 + Chr(13) + "ainda se encontram em andamento !"
      bsShowMessage(S1, "I")
      Exit Sub
    End If

    If CurrentQuery.FieldByName("GERACAODATA").IsNull Then
      bsShowMessage("Não existe informação de avaliação para esta avaliação!", "I")
      Exit Sub
    End If

    q1.Active = False
    q1.Clear
    q1.Add("SELECT SUM(PA.TOTAL) TOTAL, SUM(PA.PONTOS) PONTOS  ")
    q1.Add("  FROM SAM_AVALIACAOREF                 A,         ")
    q1.Add("       SAM_PROCREF_PRESTADOR_AVAL      PA          ")
    q1.Add(" WHERE PA.AVALIACAOREF = A.HANDLE                  ")
    q1.Add("   AND PA.PROCREFPRESTADOR = :HANDLE               ")
    q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    q1.Active = True


    CurrentQuery.Edit
    CurrentQuery.FieldByName("PONTOS").Value = q1.FieldByName("PONTOS").Value
    CurrentQuery.FieldByName("PERCENTUAL").Value = Round((q1.FieldByName("PONTOS").AsCurrency * 100) / q1.FieldByName("TOTAL").AsCurrency, 2)
    CurrentQuery.FieldByName("FINALIZACAODATA").Value = ServerNow
    CurrentQuery.FieldByName("FINALIZACAOUSUARIO").Value = CurrentUser

    'aprovação
    q1.Active = False
    q1.Clear
    q1.Add("SELECT count(1) NREC  ")
    q1.Add("  FROM SAM_PROCREF_PRESTADOR_AVAL      PA       ")
    q1.Add(" WHERE PA.PROCREFPRESTADOR = :HANDLE            ")
    q1.Add("   AND SITUACAO = 'R'                           ")
    q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    q1.Active = True

    If q1.FieldByName("NREC").AsInteger > 0 Then
      CurrentQuery.FieldByName("SITUACAO").AsString = "R"
      CurrentQuery.FieldByName("CLASSIFICACAO").AsInteger = 0
    Else
      q1.Clear
      q1.Add("SELECT COUNT(1) QTDE")
      q1.Add("  FROM SAM_PROCREF_PRESTADOR_ITEM_RES A")
      q1.Add("  JOIN SAM_PROCREF_PRESTADOR_AVAL_ITE B ON (A.PROCPRESTAVALITEM = B.HANDLE)")
      q1.Add("  JOIN SAM_PROCREF_PRESTADOR_AVAL C ON (C.HANDLE = B.PROCESSOREFPRESTAVAL)")
      q1.Add("  JOIN SAM_PROCREF_PRESTADOR D ON (D.HANDLE = C.PROCREFPRESTADOR)")
      q1.Add(" WHERE A.ELIMINATORIA = 'S' AND A.MARCA = 'S'")
      q1.Add("   AND D.HANDLE = :PHANDLE")
      q1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      q1.Active = True

      If q1.FieldByName("QTDE").AsInteger > 0 Then
        CurrentQuery.FieldByName("SITUACAO").AsString = "R"
        CurrentQuery.FieldByName("CLASSIFICACAO").AsInteger = 0
      Else
        CurrentQuery.FieldByName("SITUACAO").AsString = "A"
      End If
    End If

    CurrentQuery.Post

    RefreshNodesWithTable("SAM_PROCREF_PRESTADOR")
  End If
End Sub

Public Sub BOTAOGERAR_OnClick()

  Dim q1 As Object

  If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then
    bsShowMessage("Registro em inserção ou edição !", "I")
    Exit Sub
  End If

  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PROCREF WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If q1.FieldByName("SITUACAO").AsString = "F" Then
    bsShowMessage("Este processo de avaliação está finalizado!", "I")
    Exit Sub
  ElseIf q1.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Este processo de avaliação está cancelado!", "I")
    Exit Sub
  End If


  Dim q2 As Object
  Dim q3 As Object
  Dim qAux As Object
  Dim qExec1 As Object
  Dim qExec2 As Object
  Dim qExec3 As Object
  Dim vProcRefPrestAval As Long
  Dim vProcRefPrestAvalItem As Long

  If CurrentQuery.FieldByName("SITUACAO").AsString = "E" Then
    If Not CurrentQuery.FieldByName("GERACAODATA").IsNull Then
        bsShowMessage("Os dados das avaliações deste prestador já foram gerados!" + Chr(13) + _
          "Para gerar novamente, antes é preciso excluir os dados - botão 'Excluir'.", "Ï")
      Exit Sub
    End If
    'AVALIACAO -----------------------------------
    q1.Active = False
    q1.Clear
    q1.Add("SELECT HANDLE, APROVACAO, REPROVACAO, OBSERVACAO                         ")
    q1.Add("  FROM SAM_AVALIACAOREF WHERE HANDLE IN (SELECT AVALIACAOREF             ")
    q1.Add("                                           FROM SAM_PROCREF_AVALIACAOREF ")
    q1.Add("                                          WHERE PROCREF = :PROCREF)      ")
    q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
    q1.Active = True
    If q1.EOF Then
        bsShowMessage("Falta cadastrar avaliações para este processo!", "I")
      Exit Sub
    End If

    Set q2 = NewQuery
    q2.Add("SELECT * FROM SAM_AVALIACAOREF_ITEM WHERE AVALIACAOREF = :HANDLE")

    Set q3 = NewQuery
    q3.Add("SELECT * FROM SAM_AVALIACAOREF_ITEM_RESPOSTA WHERE AVALIACAOREFITEM = :HANDLE")

    Set qExec1 = NewQuery
    qExec1.Add("INSERT INTO SAM_PROCREF_PRESTADOR_AVAL")
    qExec1.Add("  (HANDLE,PROCREFPRESTADOR,AVALIACAOREF,TOTAL,PONTOS,PERCENTUAL,SITUACAO,REPROVACAO,OBSERVACAO) ")
    qExec1.Add(" VALUES ")
    qExec1.Add("  (:HANDLE,:PROCREFPRESTADOR,:AVALIACAOREF,:TOTAL,:PONTOS,:PERCENTUAL,:SITUACAO,:REPROVACAO,:OBSERVACAO) ")

    Set qExec2 = NewQuery
    qExec2.Add("INSERT INTO SAM_PROCREF_PRESTADOR_AVAL_ITE")
    qExec2.Add("  (HANDLE,PROCESSOREFPRESTAVAL,ORDEM,DESCRICAO,TIPO,PONTUACAOMAXIMA, PONTOS) ")
    qExec2.Add(" VALUES ")
    qExec2.Add("  (:HANDLE,:PROCESSOREFPRESTAVAL,:ORDEM,:DESCRICAO,:TIPO,:PONTUACAOMAXIMA,:PONTOS) ")

    Set qExec3 = NewQuery
    qExec3.Add("INSERT INTO SAM_PROCREF_PRESTADOR_ITEM_RES")
    qExec3.Add("  (HANDLE,PROCPRESTAVALITEM,ORDEM,DESCRICAO,DESCRITIVA,DESCRITIVATEXTO,DESCRITIVAOBRIGATORIA,PONTOSPADRAO,PONTOSPADRAOORIGINAL,PERMITEALTERARPONTO,ELIMINATORIA,MARCA,PONTOS) ")
    qExec3.Add(" VALUES ")
    qExec3.Add("  (:HANDLE,:PROCPRESTAVALITEM,:ORDEM,:DESCRICAO,:DESCRITIVA,:DESCRITIVATEXTO,:DESCRITIVAOBRIGATORIA,:PONTOSPADRAO,:PONTOSPADRAOORIGINAL,:PERMITEALTERARPONTO,:ELIMINATORIA,:MARCA,:PONTOS) ")

    Set qAux = NewQuery
    qAux.Add("SELECT SUM(PONTUACAOMAXIMA) TOTAL FROM SAM_AVALIACAOREF_ITEM WHERE AVALIACAOREF = :HANDLE")

    While Not q1.EOF
      vProcRefPrestAval = NewHandle("SAM_PROCREF_PRESTADOR_AVAL")

      qAux.Active = False
      qAux.ParamByName("HANDLE").Value = q1.FieldByName("HANDLE").Value
      qAux.Active = True

      qExec1.Active = False
      qExec1.ParamByName("HANDLE").Value = vProcRefPrestAval
      qExec1.ParamByName("PROCREFPRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").Value
      qExec1.ParamByName("AVALIACAOREF").Value = q1.FieldByName("HANDLE").Value
      qExec1.ParamByName("TOTAL").Value = qAux.FieldByName("TOTAL").Value
      qExec1.ParamByName("PONTOS").Value = 0
      qExec1.ParamByName("PERCENTUAL").Value = 0
      qExec1.ParamByName("SITUACAO").Value = "E"
      qExec1.ParamByName("REPROVACAO").Value = q1.FieldByName("REPROVACAO").Value
      qExec1.ParamByName("OBSERVACAO").Value = RTrim(q1.FieldByName("OBSERVACAO").AsString)
      qExec1.ExecSQL
      q2.Active = False
      q2.ParamByName("HANDLE").Value = q1.FieldByName("HANDLE").Value
      q2.Active = True
      While Not q2.EOF
        vProcRefPrestAvalItem = NewHandle("SAM_PROCREF_PRESTADOR_AVAL_ITE")
        qExec2.Active = False
        qExec2.ParamByName("HANDLE").Value = vProcRefPrestAvalItem
        qExec2.ParamByName("PROCESSOREFPRESTAVAL").Value = vProcRefPrestAval
        qExec2.ParamByName("ORDEM").Value = q2.FieldByName("ORDEM").Value
        qExec2.ParamByName("DESCRICAO").Value = q2.FieldByName("DESCRICAO").Value
        qExec2.ParamByName("TIPO").Value = q2.FieldByName("TIPO").Value
        qExec2.ParamByName("PONTUACAOMAXIMA").Value = q2.FieldByName("PONTUACAOMAXIMA").Value
        qExec2.ParamByName("PONTOS").Value = 0
        qExec2.ExecSQL
        q3.Active = False
        q3.ParamByName("HANDLE").Value = q2.FieldByName("HANDLE").Value
        q3.Active = True
        While Not q3.EOF
          qExec3.Active = False
          qExec3.ParamByName("HANDLE").Value = NewHandle("SAM_PROCREF_PRESTADOR_ITEM_RES")
          qExec3.ParamByName("PROCPRESTAVALITEM").Value = vProcRefPrestAvalItem
          qExec3.ParamByName("ORDEM").Value = q3.FieldByName("ORDEM").Value
          qExec3.ParamByName("DESCRICAO").Value = q3.FieldByName("DESCRICAO").Value
          qExec3.ParamByName("DESCRITIVA").Value = q3.FieldByName("DESCRITIVA").Value
          qExec3.ParamByName("DESCRITIVATEXTO").Value = q3.FieldByName("DESCRITIVATEXTO").AsString
          qExec3.ParamByName("DESCRITIVAOBRIGATORIA").Value = q3.FieldByName("DESCRITIVAOBRIGATORIA").Value
          qExec3.ParamByName("PONTOSPADRAO").Value = q3.FieldByName("PONTOS").Value
          qExec3.ParamByName("PONTOSPADRAOORIGINAL").Value = q3.FieldByName("PONTOS").Value
          qExec3.ParamByName("PERMITEALTERARPONTO").Value = q3.FieldByName("PERMITEALTERARPONTO").Value
          qExec3.ParamByName("ELIMINATORIA").Value = q3.FieldByName("ELIMINATORIA").Value
          qExec3.ParamByName("MARCA").Value = "N"
          qExec3.ParamByName("PONTOS").Value = 0
          qExec3.ExecSQL
          q3.Next
        Wend
        q2.Next
      Wend
      q1.Next
    Wend

    bsShowMessage("Somente será(ão) gerada(s) a(s) especialidade(s) para este prestador, em que o mesma não tenha sido referenciadas em outra avaliação, dentro do prazo definido no campo 'Período para nova avaliação' dos parametros gerais ", "I")

    q1.Active = False
    q1.Clear


	Dim vData As String
	vData = SQLDate( ServerDate)

	q1.Add("SELECT X.ESPECIALIDADE                                                                          ")
	q1.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE X                                                            ")
	q1.Add(" WHERE X.ESPECIALIDADE IN (SELECT P.ESPECIALIDADE                                               ")
	q1.Add("                             FROM SAM_PROCREF_ESPECIALIDADE P                                   ")
	q1.Add("                            WHERE P.PROCREF = :PROCREF)                                         ")
	q1.Add("   AND NOT EXISTS (SELECT PRE.ESPECIALIDADE                                                     ")
	q1.Add("                     FROM SAM_PROCREF_PRESTADOR_ESPEC PRE,                                      ")
	q1.Add("                          SAM_PROCREF_PRESTADOR PR                                              ")
	q1.Add("                    WHERE PRE.PROCREFPRESTADOR = PR.HANDLE                                      ")
	q1.Add("                      AND PRE.ESPECIALIDADE = X.ESPECIALIDADE                                   ")
	q1.Add("                      AND X.PRESTADOR = PR.PRESTADOR                                            ")
	q1.Add("                      AND PR.SITUACAO = 'R'                                                     ")
	q1.Add("                      AND (PR.FINALIZACAODATA IS NOT NULL                                       ")
	q1.Add("                      AND (" + SQLDateDiff(vData,"PR.FINALIZACAODATA") + "                      ")
	q1.Add("                           <= (Select PERIODONOVAAVALIACAO FROM SAM_PARAMETROSPRESTADOR)        ")
	q1.Add("                   		  ))                                                                    ")
	q1.Add("                  )                                                                             ")
	q1.Add("   AND X.DATAINICIAL <= " + vData + "                                                           ")
	q1.Add("   AND ( X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vData + " )                                  ")
	q1.Add("   AND X.PRESTADOR = :PRESTADOR                                                                 ")

	q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
	q1.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value

    q1.Active = True


    qExec1.Active = False
    qExec1.Clear
    qExec1.Add("INSERT INTO SAM_PROCREF_PRESTADOR_ESPEC      ")
    qExec1.Add("  (HANDLE,PROCREFPRESTADOR,ESPECIALIDADE)    ")
    qExec1.Add(" VALUES                                      ")
    qExec1.Add("  (:HANDLE,:PROCREFPRESTADOR,:ESPECIALIDADE) ")


    If q1.EOF Then
        bsShowMessage("Este prestador não tem as especialidades contida na avaliação." + Chr(13) + "Ou todas as especialidades do prestador já foram avaliadas em outro processo e esta com a situacao Reprovada !", "I")
      Exit Sub
    End If

    While Not q1.EOF
      qExec1.Active = False
      qExec1.ParamByName("HANDLE").Value = NewHandle("SAM_PROCREF_PRESTADOR_ESPEC")
      qExec1.ParamByName("PROCREFPRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").Value
      qExec1.ParamByName("ESPECIALIDADE").Value = q1.FieldByName("ESPECIALIDADE").Value
      qExec1.ExecSQL
      q1.Next
    Wend

    CurrentQuery.Edit
    CurrentQuery.FieldByName("GERACAODATA").Value = ServerNow
    CurrentQuery.FieldByName("GERACAOUSUARIO").Value = CurrentUser
    CurrentQuery.Post


     bsShowMessage("Processo concluído !", "I")

    '-------------------------------------------
  Else
    bsShowMessage("A avaliação deste prestador não está em andamento !", "I")
    Exit Sub
  End If

End Sub

Public Sub BOTAOREFERENCIAR_OnClick()
  Dim q1 As Object
  Dim qExec As Object
  Dim vHandle As Long
  Dim vData As Date

  If CurrentQuery.FieldByName("SITUACAO").AsString <> "A" Then
    bsShowMessage("Somente pode Referenciar um Prestador quando a Situação da Avaliação estiver Referenciável !", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
    bsShowMessage("Para referenciar um prestador, informe a Data inicial!", "I")
    Exit Sub
  End If

  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PROCREF WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If q1.FieldByName("SITUACAO").AsString = "F" Then
    bsShowMessage("Este processo de avaliação está fechado!", "I")
    Exit Sub
  ElseIf q1.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Este processo de avaliação está cancelado!", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("FINALIZACAODATA").IsNull Then
    bsShowMessage("Esta avaliação ainda não foi finalizada !", "I")
    Exit Sub
  End If


  If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then
    bsShowMessage("Registro em inserção ou edição !", "I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("REFERENCIAMENTODATA").IsNull Then
    bsShowMessage("Este prestador já está referenciado nesta avaliação!", "I")
    Exit Sub
  End If

  Set qParametro = NewQuery
  qParametro.Active = False
  qParametro.Clear
  qParametro.Add("SELECT TEMPODISPPRESTADOR, TEMPOREFERENCIAMENTO FROM SAM_PARAMETROSPRESTADOR")
  qParametro.Active = True
  vData = DateAdd("m", qParametro.FieldByName("TEMPODISPPRESTADOR").AsInteger, CurrentQuery.FieldByName("FINALIZACAODATA").Value)

  If ServerDate > CDate(vData) Then
    bsShowMessage("Este prestador precisa passar por uma nova avaliação !", "I")
    Exit Sub
  End If

  CurrentQuery.Edit
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CurrentQuery.FieldByName("DATAFINAL").Value = DateAdd("m", qParametro.FieldByName("TEMPOREFERENCIAMENTO").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").Value)
  End If
  CurrentQuery.FieldByName("REFERENCIAMENTODATA").Value = ServerNow
  CurrentQuery.FieldByName("REFERENCIAMENTOUSUARIO").AsInteger = CurrentUser
  CurrentQuery.FieldByName("SITUACAO").AsString = "M"

  Set qParametro = Nothing

  q1.Active = True
  q1.Clear
  q1.Add("SELECT COUNT(HANDLE) NREC FROM SAM_PROCREF_PRESTADOR WHERE :DATA BETWEEN DATAINICIAL AND DATAFINAL AND PONTOS > :PONTOS AND HANDLE <> :HANDLE")
  q1.ParamByName("DATA").Value = ServerDate
  q1.ParamByName("PONTOS").Value = CurrentQuery.FieldByName("PONTOS").AsInteger
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.Active = True

  CurrentQuery.FieldByName("CLASSIFICACAO").Value = q1.FieldByName("NREC").AsInteger + 1

  CurrentQuery.Post

  q1.Active = False
  q1.Clear
  q1.Add("UPDATE SAM_PROCREF_PRESTADOR                                                              ")
  q1.Add("  SET CLASSIFICACAO = CLASSIFICACAO + 1                                                   ")
  q1.Add("WHERE :DATA BETWEEN DATAINICIAL AND DATAFINAL AND PONTOS <= :PONTOS AND HANDLE <> :HANDLE ")
  q1.ParamByName("DATA").Value = ServerDate
  q1.ParamByName("PONTOS").Value = CurrentQuery.FieldByName("PONTOS").AsInteger
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.ExecSQL

  Set qExec = NewQuery
  qExec.Add("INSERT INTO SAM_PRESTADOR_REFERENCIA                                             ")
  If Not CurrentQuery.FieldByName("OBSERVACAO").IsNull Then
    qExec.Add("  (HANDLE,PRESTADOR,DATAINICIAL,DATAFINAL,MOTIVOREFERENCIAMENTO,OBSERVACAO, PROCREFPRESTADOR, CLASSIFICACAO)         ")
    qExec.Add("VALUES                                                                           ")
    qExec.Add("  (:HANDLE,:PRESTADOR,:DATAINICIAL,:DATAFINAL,:MOTIVOREFERENCIAMENTO,:OBSERVACAO, :PROCREFPRESTADOR, :CLASSIFICACAO) ")
  Else
    qExec.Add("  (HANDLE,PRESTADOR,DATAINICIAL,DATAFINAL,MOTIVOREFERENCIAMENTO, PROCREFPRESTADOR, CLASSIFICACAO)         ")
    qExec.Add("VALUES                                                                           ")
    qExec.Add("  (:HANDLE,:PRESTADOR,:DATAINICIAL,:DATAFINAL,:MOTIVOREFERENCIAMENTO, :PROCREFPRESTADOR, :CLASSIFICACAO) ")
  End If

  vHandle = NewHandle("SAM_PRESTADOR_REFERENCIA")
  qExec.ParamByName("HANDLE").Value = vHandle
  qExec.ParamByName("PROCREFPRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").Value
  qExec.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  qExec.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  qExec.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  qExec.ParamByName("MOTIVOREFERENCIAMENTO").Value = CurrentQuery.FieldByName("MOTIVOREFERENCIAMENTO").Value
  If Not CurrentQuery.FieldByName("OBSERVACAO").IsNull Then
    qExec.ParamByName("OBSERVACAO").Value = CurrentQuery.FieldByName("OBSERVACAO").Value
  End If
  qExec.ParamByName("CLASSIFICACAO").Value = CurrentQuery.FieldByName("CLASSIFICACAO").Value
  qExec.ExecSQL

  qExec.Active = False
  qExec.Clear
  qExec.Add("INSERT INTO SAM_PRESTADOR_REFERENCIA_ESPEC         ")
  qExec.Add("  (HANDLE,PRESTADORREFERENCIA,ESPECIALIDADE)       ")
  qExec.Add("VALUES                                             ")
  qExec.Add("  (:HANDLE,:PRESTADORREFERENCIA,:ESPECIALIDADE)    ")

  q1.Active = False
  q1.Clear
  q1.Add("SELECT ESPECIALIDADE FROM SAM_PROCREF_PRESTADOR_ESPEC WHERE PROCREFPRESTADOR = :PROCREFPRESTADOR")
  q1.ParamByName("PROCREFPRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.Active = True
  While Not q1.EOF
    qExec.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_REFERENCIA_ESPEC")
    qExec.ParamByName("PRESTADORREFERENCIA").Value = vHandle
    qExec.ParamByName("ESPECIALIDADE").Value = q1.FieldByName("ESPECIALIDADE").Value
    qExec.ExecSQL
    q1.Next
  Wend


  'SMS 41006 - Milani - 13/06/2005

  '  If Not CurrentQuery.FieldByName("ESPECIALIDADE").IsNull Then

  '    qExec.Acti e=False
  '    qExec.Clear
  '    qExec.Add("INSERT INTO SAM_PRESTADOR_ESPECIALIDADE                                                                              ")
  '    qExec.Add("  (HANDLE,PRESTADOR,ESPECIALIDADE,DATAINICIAL,PRINCIPAL, TEMPORARIO,HERDOUDOGRUPOEMPRESARIAL,PUBLICARNOLIVRO, PUBLICARINTERNET,VISUALIZARCENTRAL)")
  '    qExec.Add("VALUES                                                                           ")
  '    qExec.Add("  (:HANDLE,:PRESTADOR,:ESPECIALIDADE,:DATAINICIAL,:PRINCIPAL,:TEMPORARIO,:HERDOUDOGRUPOEMPRESARIAL,:PUBLICARNOLIVRO, :PUBLICARINTERNET,:VISUALIZARCENTRAL) ")

  '    vHandle = NewHandle("SAM_PRESTADOR_ESPECIALIDADE")
  '    qExec.ParamByName("HANDLE").Value         = vHandle
  '    qExec.ParamByName("ESPECIALIDADE").Value  = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  '    qExec.ParamByName("PRESTADOR").Value      = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  '    qExec.ParamByName("DATAINICIAL").Value    = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime

  '    qExec.ParamByName("PRINCIPAL").Value                = "N"
  '    qExec.ParamByName("TEMPORARIO").Value               = "N"
  '    qExec.ParamByName("HERDOUDOGRUPOEMPRESARIAL").Value = "N"
  '    qExec.ParamByName("PUBLICARNOLIVRO").Value          = "N"
  '    qExec.ParamByName("PUBLICARINTERNET").Value         = "N"
  '    qExec.ParamByName("VISUALIZARCENTRAL").Value        = "N"
  '    qExec.ExecSQL

  '  End If

  CurrentQuery.Active = False
  CurrentQuery.Active = True


End Sub

Public Sub DATAINICIAL_OnExit()

  If (Not CurrentQuery.FieldByName("DATAINICIAL").IsNull) And (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then

    Set qParametro = NewQuery
    qParametro.Active = False
    qParametro.Clear
    qParametro.Add("SELECT TEMPODISPPRESTADOR, TEMPOREFERENCIAMENTO FROM SAM_PARAMETROSPRESTADOR")
    qParametro.Active = True

    If CurrentQuery.State = 3 Or CurrentQuery.State = 2 Then
      CurrentQuery.FieldByName("DATAFINAL").Value = DateAdd("m", qParametro.FieldByName("TEMPOREFERENCIAMENTO").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").Value)
    End If
  End If

  Set qParametro = Nothing

End Sub


Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim q1 As Object
  Dim q2 As Object
  Dim Interface As Object
  Set q1 = NewQuery
  q1.Add("SELECT COUNT(HANDLE) NREC FROM SAM_PROCREF_ESPECIALIDADE WHERE PROCREF = :PROCREF")
  q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If q1.FieldByName("NREC").AsInteger = 0 Then
    If VisibleMode Then
      MsgBox "Falta informa especialidades para este processo!"
    End If
    Exit Sub
  End If

  Set q2 = NewQuery
  q2.Add("SELECT FILIAL, MUNICIPIO FROM SAM_PROCREF WHERE HANDLE = " + CurrentQuery.FieldByName("PROCREF").AsString)
  q2.Active = True


  If VisibleMode Then
    Dim vPrestador As Long
    Dim vData As String
    Dim vCriterio As String

    ShowPopup = False

    vData = SQLDate(ServerDate)

    Set Interface = CreateBennerObject("Procura.Procurar")

    vCriterio = "SAM_PRESTADOR.HANDLE IN (SELECT X.PRESTADOR " + _
                "                           FROM SAM_PRESTADOR_ESPECIALIDADE X " + _
                "                          WHERE X.ESPECIALIDADE IN (SELECT P.ESPECIALIDADE" + _
                "                                                      FROM SAM_PROCREF_ESPECIALIDADE P" + _
                "                                                     WHERE P.PROCREF = " + CurrentQuery.FieldByName("PROCREF").AsString + _
                "                                                   ) " + _
                "                            AND X.DATAINICIAL <= " + vData + _
                "                            AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vData + ")" + _
                "                        )" + _
                " AND SAM_PRESTADOR.DATACREDENCIAMENTO IS NOT NULL " + _
                " AND (SAM_PRESTADOR.DATADESCREDENCIAMENTO IS NULL OR SAM_PRESTADOR.DATADESCREDENCIAMENTO > " + vData + ")" + _
                " AND SAM_PRESTADOR.FILIALPADRAO = " + q2.FieldByName("FILIAL").AsString
    If Not q2.FieldByName("MUNICIPIO").IsNull Then
      vCriterio = vCriterio + " AND SAM_PRESTADOR.MUNICIPIOPAGAMENTO = " + q2.FieldByName("MUNICIPIO").AsString
    End If

    vPrestador = Interface.Exec(CurrentSystem, "SAM_PRESTADOR|", "PRESTADOR|NOME|CPFCNPJ", 3, "Prestador|Nome|CPF/CNPJ", vCriterio, "Tabela de prestadores", False, "", "")

    If vPrestador > 0 Then
      CurrentQuery.FieldByName("PRESTADOR").Value = vPrestador
    End If
  End If

End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("INCLUSAODATA").Value = ServerNow
  CurrentQuery.FieldByName("INCLUSAOUSUARIO").Value = CurrentUser
End Sub

Public Sub TABLE_AfterPost()

  Dim q1 As Object
  Dim qExec As Object

  vBotaoAprovar = "N"

  Set q1 = NewQuery
  q1.Add("SELECT X.APR, Y.REP, Z.TOT             ")
  q1.Add("  FROM (SELECT COUNT(HANDLE) APR       ")
  q1.Add("          FROM SAM_PROCREF_PRESTADOR   ")
  q1.Add("         WHERE SITUACAO = 'A'          ")
  q1.Add("           AND PROCREF = :PROCREF      ")
  q1.Add("       ) X,                            ")
  q1.Add("       (SELECT COUNT(HANDLE) REP       ")
  q1.Add("          FROM SAM_PROCREF_PRESTADOR   ")
  q1.Add("         WHERE SITUACAO = 'R'          ")
  q1.Add("           AND PROCREF = :PROCREF      ")
  q1.Add("       ) Y,                            ")
  q1.Add("       (SELECT COUNT(HANDLE) TOT       ")
  q1.Add("          FROM SAM_PROCREF_PRESTADOR   ")
  q1.Add("         WHERE PROCREF = :PROCREF      ")
  q1.Add("       ) Z                             ")
  q1.ParamByName("PROCREF").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  Set qExec = NewQuery
  qExec.Add("UPDATE SAM_PROCREF                   ")
  qExec.Add("  SET QTDETOTAL = :QTDETOTAL,        ")
  qExec.Add("      QTDEAPROVADO = :QTDEAPROVADO,  ")
  qExec.Add("      QTDEREPROVADO = :QTDEREPROVADO ")
  qExec.Add("WHERE HANDLE = :HANDLE               ")
  qExec.ParamByName("QTDETOTAL").Value = q1.FieldByName("TOT").Value
  qExec.ParamByName("QTDEAPROVADO").Value = q1.FieldByName("APR").Value
  qExec.ParamByName("QTDEREPROVADO").Value = q1.FieldByName("REP").Value
  qExec.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  qExec.ExecSQL


End Sub

Public Sub TABLE_AfterScroll()


  Dim q2 As Object

  Set q2 = NewQuery
  q2.Add("SELECT FILIAL, MUNICIPIO FROM SAM_PROCREF WHERE HANDLE = " + CStr(RecordHandleOfTable("SAM_PROCREF")))
  q2.Active = True

  Dim vData As String
  Dim vAux As String

  vData = SQLDate(ServerDate)

  If WebMode Then


    vAux = "@ALIAS.HANDLE IN (SELECT X.PRESTADOR " + _
                "                           FROM SAM_PRESTADOR_ESPECIALIDADE X " + _
                "                          WHERE X.ESPECIALIDADE IN (SELECT P.ESPECIALIDADE" + _
                "                                                      FROM SAM_PROCREF_ESPECIALIDADE P" + _
                "                                                     WHERE P.PROCREF = @CAMPO(PROCREF)" + _
                "                                                   ) " + _
                "                            AND X.DATAINICIAL <= " + vData + _
                "                            AND (X.DATAFINAL IS NULL OR X.DATAFINAL >= " + vData + ")" + _
                "                        )" + _
                " AND @ALIAS.DATACREDENCIAMENTO IS NOT NULL " + _
                " AND (@ALIAS.DATADESCREDENCIAMENTO IS NULL OR @ALIAS.DATADESCREDENCIAMENTO > " + vData + ")" + _
                " AND @ALIAS.FILIALPADRAO = " + q2.FieldByName("FILIAL").AsString
    If Not q2.FieldByName("MUNICIPIO").IsNull Then
      vAux = vAux + " AND @ALIAS.MUNICIPIOPAGAMENTO = " + q2.FieldByName("MUNICIPIO").AsString
    End If

	PRESTADOR.WebLocalWhere = vAux

  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "R" Then
    BOTAOAPROVAR.Enabled = True 'se o prestador estiver reprovado, habilitará o botão para Aprovar.
  Else
    BOTAOAPROVAR.Enabled = False
  End If

  APROVACAOJUSTIFICATIVA.ReadOnly = True

  If CurrentQuery.FieldByName("SITUACAO").AsString = "R" Or CurrentQuery.FieldByName("SITUACAO").AsString = "A" Or CurrentQuery.FieldByName("SITUACAO").AsString = "M" Then
    PRESTADOR.ReadOnly = True
    MOTIVOREFERENCIAMENTO.ReadOnly = True
    OBSERVACAO.ReadOnly = True
  Else
    PRESTADOR.ReadOnly = False
    MOTIVOREFERENCIAMENTO.ReadOnly = False
    OBSERVACAO.ReadOnly = False
  End If

  If Not CurrentQuery.FieldByName("REFERENCIAMENTODATA").IsNull Then
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
  Else
    DATAINICIAL.ReadOnly = False
    DATAFINAL.ReadOnly = False
  End If

End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

	If WebMode Then
  		ESTADO.WebLocalWhere = "A.HANDLE IN (SELECT ESTADO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @CAMPO(PRESTADOR))"
	ElseIf VisibleMode Then
  		ESTADO.LocalWhere = "ESTADOS.HANDLE IN (SELECT ESTADO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @PRESTADOR)"
	End If

    If WebMode Then
    	MUNICIPIO.WebLocalWhere = "MUNICIPIOS.HANDLE IN (SELECT MUNICIPIO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @CAMPO(PRESTADOR))"
    ElseIf VisibleMode Then
    	MUNICIPIO.LocalWhere = "MUNICIPIOS.HANDLE IN (SELECT MUNICIPIO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @PRESTADOR)"
    End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
  		ESTADO.WebLocalWhere = "A.HANDLE IN (SELECT ESTADO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @CAMPO(PRESTADOR))"
	ElseIf VisibleMode Then
  		ESTADO.LocalWhere = "ESTADOS.HANDLE IN (SELECT ESTADO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @PRESTADOR)"
	End If

    If WebMode Then
    	MUNICIPIO.WebLocalWhere = "MUNICIPIOS.HANDLE IN (SELECT MUNICIPIO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @CAMPO(PRESTADOR))"
    ElseIf VisibleMode Then
    	MUNICIPIO.LocalWhere = "MUNICIPIOS.HANDLE IN (SELECT MUNICIPIO FROM SAM_PRESTADOR_ENDERECO WHERE PRESTADOR = @PRESTADOR)"
    End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim q1 As Object
  Dim q2 As Object
  Dim qAux As Object

  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO, MUNICIPIO FROM SAM_PROCREF WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCREF").Value
  q1.Active = True

  If Not q1.FieldByName("MUNICIPIO").IsNull Then
    Set q2 = NewQuery
    q2.Add("SELECT ESTADO FROM MUNICIPIOS WHERE HANDLE = " + q1.FieldByName("MUNICIPIO").AsString)
    q2.Active = True
    CurrentQuery.FieldByName("MUNICIPIO").Value = q1.FieldByName("MUNICIPIO").Value
    CurrentQuery.FieldByName("ESTADO").Value = q2.FieldByName("ESTADO").Value
  End If

  CurrentQuery.FieldByName("ALTERACAODATA").Value = ServerNow
  CurrentQuery.FieldByName("ALTERACAOUSUARIO").Value = CurrentUser
  If vBotaoAprovar = "S" And RTrim(LTrim(CurrentQuery.FieldByName("APROVACAOJUSTIFICATIVA").AsString)) = "" Then
    bsShowMessage("Para fazer aprovação deve ser informado o campo 'Justificativa' ", "E")
    CanContinue = False
    Exit Sub
  End If


  If q1.FieldByName("SITUACAO").AsString = "F" Then
    bsShowMessage("Este processo de avaliação está finalizado!", "E")
    CanContinue = False
    Exit Sub
  ElseIf q1.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Este processo de avaliação está cancelado!", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOAPROVAR"
			BOTAOAPROVAR_OnClick
		Case "BOTAOEXCLUIR"
			BOTAOEXCLUIR_OnClick
		Case "BOTAOFINALIZAR"
			BOTAOFINALIZAR_OnClick
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOREFERENCIAR"
			BOTAOREFERENCIAR_OnClick
	End Select
End Sub
