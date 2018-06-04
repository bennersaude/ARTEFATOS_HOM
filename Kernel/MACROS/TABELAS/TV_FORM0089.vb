'HASH: 58C06E04D9393412FC0D47CA72CB9557
'#Uses "*bsShowMessage"

Public Sub TABACAO_OnChange()
  If (TABACAO.PageIndex = 0) Then 'Inclusão
    TABINCLUSAO_OnChange
  ElseIf (TABACAO.PageIndex = 1) Then 'Alteração
    TABALTERACAO_OnChange
  ElseIf (TABACAO.PageIndex = 2) Then 'Exclusão
    TABEXCLUSAO_OnChange
  ElseIf (TABACAO.PageIndex = 3) Then 'Vinculação
    TABVINCULACAO_OnChange
  End If
End Sub

Public Sub TABINCLUSAO_OnChange()
  If TABINCLUSAO.PageIndex = 1 Then 'Endereços
    ENDERECOS.WebLocalWhere = " (DATACANCELAMENTO IS NULL OR (DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO > " + SQLDate(ServerDate) + "))"
    ENDERECOS.Where = " (DATACANCELAMENTO IS NULL OR (DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO > " + SQLDate(ServerDate) + "))"
  End If
End Sub

Public Sub TABALTERACAO_OnChange()
  If TABALTERACAO.PageIndex = 1 Then 'Endereços
    ENDERECOS.WebLocalWhere = " (DATACANCELAMENTO IS NULL OR (DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO > " + SQLDate(ServerDate) + "))"
    ENDERECOS.Where = " (DATACANCELAMENTO IS NULL OR (DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO > " + SQLDate(ServerDate) + "))"
  End If
End Sub

Public Sub TABEXCLUSAO_OnChange()
  If TABEXCLUSAO.PageIndex = 1 Then 'Endereços
    ENDERECOS.WebLocalWhere = " DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO <= " + SQLDate(ServerDate)
    ENDERECOS.Where = " DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO <= " + SQLDate(ServerDate) 
  End If
End Sub

Public Sub TABVINCULACAO_OnChange()
  If TABVINCULACAO.PageIndex = 1 Then 'Endereços
    ENDERECOS.WebLocalWhere = " (DATACANCELAMENTO IS NULL OR (DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO > " + SQLDate(ServerDate) + "))"
    ENDERECOS.Where = " (DATACANCELAMENTO IS NULL OR (DATACANCELAMENTO IS NOT NULL AND DATACANCELAMENTO > " + SQLDate(ServerDate) + "))"
  End If
End Sub

Public Sub TABLE_AfterScroll()
  gTabAcao = 1
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim handleRotina As Integer
  handleRotina = RecordHandleOfTable("ANS_ROTINAENVIORPS")

  Dim qVerificaRotina As Object
  Set qVerificaRotina = NewQuery
  qVerificaRotina.Clear
  qVerificaRotina.Add("SELECT SITUACAOEXCLUSAODADOS ")
  qVerificaRotina.Add("  FROM ANS_ROTINAENVIORPS    ")
  qVerificaRotina.Add(" WHERE HANDLE = :HANDLE      ")
  qVerificaRotina.ParamByName("HANDLE").AsInteger = handleRotina
  qVerificaRotina.Active = True

  If (qVerificaRotina.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 1) And (qVerificaRotina.FieldByName("SITUACAOEXCLUSAODADOS").AsInteger <> 5)Then
    bsShowMessage("Dados não podem ser gerados, pois ainda existem registros a serem excluídos", "I")
    Exit Sub
  End If

  Set qVerificaRotina = Nothing


  Dim sx As CSServerExec
  Set sx = NewServerExec
  Dim DadosRotina As String

  If CurrentQuery.FieldByName("TABACAO").AsInteger = 1 Then
    DadosRotina = "Inclusão"
    If CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 1 Then
      DadosRotina = DadosRotina + " - Tipo Prestadores"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 2 Then
      DadosRotina = DadosRotina + " - Tipo Prestador x Endereços"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 3 Then
      DadosRotina = DadosRotina + " - Tipo Vigência"
    End If
  ElseIf CurrentQuery.FieldByName("TABACAO").AsInteger = 2 Then
    DadosRotina = "Alteração"
    If CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 1 Then
      DadosRotina = DadosRotina + " - Tipo Prestadores"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 2 Then
      DadosRotina = DadosRotina + " - Tipo Prestador x Endereços"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 3 Then
      DadosRotina = DadosRotina + " - Tipo Vigência"
    End If
  ElseIf CurrentQuery.FieldByName("TABACAO").AsInteger = 3 Then
    DadosRotina = "Exclusão"
    If CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 1 Then
      DadosRotina = DadosRotina + " - Tipo Prestadores"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 2 Then
      DadosRotina = DadosRotina + " - Tipo Prestador x Endereços"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 3 Then
      DadosRotina = DadosRotina + " - Tipo Vigência"
    End If
  ElseIf CurrentQuery.FieldByName("TABACAO").AsInteger = 4 Then
    DadosRotina = "Vinculação"
    If CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 1 Then
      DadosRotina = DadosRotina + " - Tipo Prestadores"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 2 Then
      DadosRotina = DadosRotina + " - Tipo Prestador x Endereços"
    ElseIf CurrentQuery.FieldByName("TABINCLUSAO").AsInteger = 3 Then
      DadosRotina = DadosRotina + " - Tipo Vigência"
    End If
  End If
  sx.Description = "RPS - Geração de dados [" + DadosRotina + "]
  sx.DllClassName = "Benner.Saude.ANS.Processos.GerarDadosRPS"

  sx.SessionVar("RPS_HANDLEROTINA") = CStr(handleRotina)
  sx.SessionVar("RPS_TABACAO") = CurrentQuery.FieldByName("TABACAO").AsString
  sx.SessionVar("RPS_TABINCLUSAO") = CurrentQuery.FieldByName("TABINCLUSAO").AsString
  sx.SessionVar("RPS_TABALTERACAO") = CurrentQuery.FieldByName("TABALTERACAO").AsString
  sx.SessionVar("RPS_TABEXCLUSAO") = CurrentQuery.FieldByName("TABEXCLUSAO").AsString
  sx.SessionVar("RPS_TABVINCULACAO") = CurrentQuery.FieldByName("TABVINCULACAO").AsString
  sx.SessionVar("RPS_PRESTADORES") = CurrentQuery.FieldByName("PRESTADORES").AsString
  sx.SessionVar("RPS_PRESTADOR") = CurrentQuery.FieldByName("PRESTADOR").AsString
  sx.SessionVar("RPS_ENDERECOS") = CurrentQuery.FieldByName("ENDERECOS").AsString
  sx.SessionVar("RPS_DATAINICIAL") = CurrentQuery.FieldByName("DATAINICIAL").AsString
  sx.SessionVar("RPS_DATAFINAL") = CurrentQuery.FieldByName("DATAFINAL").AsString

  sx.Execute
  Set sx = Nothing

  bsShowMessage("Processo enviado para o servidor ", "I")

End Sub
