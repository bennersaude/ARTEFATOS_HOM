'HASH: E34224896C1D2819B34B60806F795799
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificacao As Object
  Set qVerificacao = NewQuery

  qVerificacao.Active=False
  qVerificacao.Clear
  qVerificacao.Add("SELECT COUNT(1) QTD                                  ")
  qVerificacao.Add("  FROM SFN_DMEDANOCALENDARIO_PROC ")
  qVerificacao.Add(" WHERE ANOCALENDARIO = :ANO                ")
  qVerificacao.ParamByName("ANO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
  qVerificacao.Active = True

  If qVerificacao.FieldByName("QTD").AsInteger = 0 Then
    CanContinue = False
    bsShowMessage("Não há processamento a ser cancelado/excluído para o ano-calendário selecionado!", "E")
    Set qVerificacao = Nothing
    Exit Sub
  End If

  Dim viStatusProcesso As Long
  Dim viStatusProcessoArq As Long
  Dim objServerExec As CSServerExec
  Dim objServerExecArq As CSServerExec

  viStatusProcesso = 0
  viStatusProcessoArq = 0

  qVerificacao.Active=False
  qVerificacao.Clear
  qVerificacao.Add("SELECT HANDLE, PROCESSO, PROCESSOARQUIVO    ")
  qVerificacao.Add("  FROM SFN_DMEDANOCALENDARIO_PROC                 ")
  qVerificacao.Add(" WHERE ANOCALENDARIO = :ANO                                ")
  qVerificacao.Add("   ORDER BY HANDLE DESC                                          ")
  qVerificacao.ParamByName("ANO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
  qVerificacao.Active = True

  If Not qVerificacao.FieldByName("PROCESSO").IsNull Then
    Set objServerExec = GetServerExec(qVerificacao.FieldByName("PROCESSO").AsInteger)
    viStatusProcesso = objServerExec.Status
  End If

  If Not qVerificacao.FieldByName("PROCESSOARQUIVO").IsNull Then
    Set objServerExecArq = GetServerExec(qVerificacao.FieldByName("PROCESSOARQUIVO").AsInteger)
    viStatusProcessoArq = objServerExecArq.Status
  End If

  If viStatusProcesso = esRunning Or viStatusProcessoArq = esRunning Then
    If bsShowMessage("Existe um processo em execução para o ano-calendário selecionado. Deseja continuar? ", "Q") = vbYes Then
      If viStatusProcesso = esRunning Then
        objServerExec.RequestAbort
      End If

      If viStatusProcessoArq = esRunning Then
        objServerExecArq.RequestAbort
      End If
      ExcluirRegistrosDmed(CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger, qVerificacao.FieldByName("HANDLE").AsInteger)
      bsShowMessage("Cancelamento efetuado com sucesso!", "I")
    End If
  Else
    ExcluirRegistrosDmed(CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger, qVerificacao.FieldByName("HANDLE").AsInteger)
    bsShowMessage("Cancelamento efetuado com sucesso!", "I")
  End If

  Set qVerificacao = Nothing
  Set objServerExec = Nothing
  Set objServerExecArq = Nothing

End Sub

Public Function ExcluirRegistrosDmed(viAnoCalendario As Long, viAnoCalendarioProc As Long)

  Dim qExclusao   As Object
  Set qExclusao = NewQuery

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_RDTOP                                                                                          ")
  qExclusao.Add(" WHERE ANOCALENDARIODMEDDTOP IN (SELECT DTOP.HANDLE                                                                         ")
  qExclusao.Add("                                   FROM SFN_DMEDANOCALENDARIO_DTOP  DTOP                                                    ")
  qExclusao.Add("                                   JOIN SFN_DMEDANOCALENDARIO_TOP   DMEDTOP ON (DMEDTOP.HANDLE = DTOP.ANOCALENDARIODMEDTOP) ")
  qExclusao.Add("                                   JOIN SFN_DMEDANOCALENDARIO       DMED    ON (DMED.HANDLE = DMEDTOP.ANOCALENDARIODMED)    ")
  qExclusao.Add("                                  WHERE DMED.HANDLE = :ANOCALENDARIO )                                                      ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = viAnoCalendario
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_DTOP                                                                                           ")
  qExclusao.Add(" WHERE ANOCALENDARIODMEDTOP IN (SELECT DMEDTOP.HANDLE                                                                       ")
  qExclusao.Add("                                  FROM SFN_DMEDANOCALENDARIO_TOP   DMEDTOP                                                  ")
  qExclusao.Add("                                  JOIN SFN_DMEDANOCALENDARIO       DMED    ON (DMED.HANDLE = DMEDTOP.ANOCALENDARIODMED)     ")
  qExclusao.Add("                                 WHERE DMED.HANDLE = :ANOCALENDARIO )                                                       ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = viAnoCalendario
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_RTOP                                                                                           ")
  qExclusao.Add(" WHERE ANOCALENDARIODMEDTOP IN (SELECT DMEDTOP.HANDLE                                                                       ")
  qExclusao.Add("                                  FROM SFN_DMEDANOCALENDARIO_TOP   DMEDTOP                                                  ")
  qExclusao.Add("                                  JOIN SFN_DMEDANOCALENDARIO       DMED    ON (DMED.HANDLE = DMEDTOP.ANOCALENDARIODMED)     ")
  qExclusao.Add("                                 WHERE DMED.HANDLE = :ANOCALENDARIO )                                                       ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = viAnoCalendario
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_TOP                                                                                            ")
  qExclusao.Add(" WHERE ANOCALENDARIODMED = :ANOCALENDARIO                                                                                   ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = viAnoCalendario
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add(" DELETE FROM SFN_DMEDOCORRENCIAS WHERE DMEDANOCALENDARIOPROC = :HANDLE                                                      ")
  qExclusao.ParamByName("HANDLE").AsInteger = viAnoCalendarioProc
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add(" DELETE FROM SFN_DMEDANOCALENDARIO_PROC WHERE HANDLE = :HANDLE                                                              ")
  qExclusao.ParamByName("HANDLE").AsInteger = viAnoCalendarioProc
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_IR_INTERNET                                                                                                      ")
  qExclusao.Add(" WHERE ANOCALENDARIO = :ANOCALENDARIO                                                                                       ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = viAnoCalendario
  qExclusao.ExecSQL

  Set qExclusao    = Nothing

End Function
