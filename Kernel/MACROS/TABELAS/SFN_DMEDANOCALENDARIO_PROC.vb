'HASH: 866255A06772E7223592D8EB3D94F319
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim qVerificacao As Object
  Dim qExclusao    As Object

  Set qVerificacao = NewQuery
  Set qExclusao    = NewQuery

  qVerificacao.Active=False
  qVerificacao.Clear
  qVerificacao.Add("SELECT COUNT(1) QTD               ")
  qVerificacao.Add("  FROM SFN_DMEDANOCALENDARIO_PROC ")
  qVerificacao.Add(" WHERE ANOCALENDARIO = :ANO       ")
  qVerificacao.Add("   AND HANDLE > :HANDLE           ")

  qVerificacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificacao.ParamByName("ANO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger

  qVerificacao.Active = True

  If qVerificacao.FieldByName("QTD").AsInteger > 0 Then
    CanContinue = False
    bsShowMessage("Somente o último processamento pode ser excluído", "E")
    Set qVerificacao = Nothing
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("PROCESSO").IsNull Then
	Dim objServerExec As CSServerExec
	Set objServerExec = GetServerExec(CurrentQuery.FieldByName("PROCESSO").AsInteger)

	If objServerExec.Status = esNone Then
	  CanContinue = False
	  bsShowMessage("O processamento ainda não foi iniciado! Verifique o monitor de processos", "E")
	  Exit Sub
	End If

	If objServerExec.Status = esRunning Then
  	  CanContinue = False
	  bsShowMessage("O processamento ainda está em progresso! Verifique o monitor de processos", "E")
	  Exit Sub
	End If

	If Not CurrentQuery.FieldByName("PROCESSOARQUIVO").IsNull Then
	  Set objServerExec = GetServerExec(CurrentQuery.FieldByName("PROCESSOARQUIVO").AsInteger)

	  If objServerExec.Status = esNone Then
	    CanContinue = False
	    bsShowMessage("A geração de arquivos foi solicitada mas ainda não foi iniciada! Verifique o monitor de processos", "E")
	    Exit Sub
      End If

	  If objServerExec.Status = esRunning Then
	    CanContinue = False
	    bsShowMessage("A geração de arquivos ainda está em progresso! Verifique o monitor de processos", "E")
	    Exit Sub
	  End If

    End If

  End If

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_RDTOP                                                                                          ")
  qExclusao.Add(" WHERE ANOCALENDARIODMEDDTOP IN (SELECT DTOP.HANDLE                                                                         ")
  qExclusao.Add("                                   FROM SFN_DMEDANOCALENDARIO_DTOP  DTOP                                                    ")
  qExclusao.Add("                                   JOIN SFN_DMEDANOCALENDARIO_TOP   DMEDTOP ON (DMEDTOP.HANDLE = DTOP.ANOCALENDARIODMEDTOP) ")
  qExclusao.Add("                                   JOIN SFN_DMEDANOCALENDARIO       DMED    ON (DMED.HANDLE = DMEDTOP.ANOCALENDARIODMED)    ")
  qExclusao.Add("                                  WHERE DMED.HANDLE = :ANOCALENDARIO )                                                      ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_DTOP                                                                                           ")
  qExclusao.Add(" WHERE ANOCALENDARIODMEDTOP IN (SELECT DMEDTOP.HANDLE                                                                       ")
  qExclusao.Add("                                  FROM SFN_DMEDANOCALENDARIO_TOP   DMEDTOP                                                  ")
  qExclusao.Add("                                  JOIN SFN_DMEDANOCALENDARIO       DMED    ON (DMED.HANDLE = DMEDTOP.ANOCALENDARIODMED)     ")
  qExclusao.Add("                                 WHERE DMED.HANDLE = :ANOCALENDARIO )                                                       ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_RTOP                                                                                           ")
  qExclusao.Add(" WHERE ANOCALENDARIODMEDTOP IN (SELECT DMEDTOP.HANDLE                                                                       ")
  qExclusao.Add("                                  FROM SFN_DMEDANOCALENDARIO_TOP   DMEDTOP                                                  ")
  qExclusao.Add("                                  JOIN SFN_DMEDANOCALENDARIO       DMED    ON (DMED.HANDLE = DMEDTOP.ANOCALENDARIODMED)     ")
  qExclusao.Add("                                 WHERE DMED.HANDLE = :ANOCALENDARIO )                                                       ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDANOCALENDARIO_TOP                                                                                            ")
  qExclusao.Add(" WHERE ANOCALENDARIODMED = :ANOCALENDARIO                                                                                   ")
  qExclusao.ParamByName("ANOCALENDARIO").AsInteger = CurrentQuery.FieldByName("ANOCALENDARIO").AsInteger
  qExclusao.ExecSQL

  qExclusao.Clear
  qExclusao.Add("DELETE                                                                                                                      ")
  qExclusao.Add("  FROM SFN_DMEDOCORRENCIAS                                                                                                  ")
  qExclusao.Add(" WHERE DMEDANOCALENDARIOPROC = :DMEDANOCALENDARIOPROC                                                                       ")
  qExclusao.ParamByName("DMEDANOCALENDARIOPROC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qExclusao.ExecSQL

  Set qExclusao    = Nothing
  RefreshNodesWithTable("SFN_DMEDANOCALENDARIO_TOP")

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("NUMERORECIBO").AsString <> "") Then
    If (CurrentQuery.FieldByName("NUMERORECIBO").Value <=0 )  Then
   	  CanContinue = False
	  bsShowMessage("O número do recibo deve ser informado com valores positivos (maior que zero)", "E")
	  Exit Sub
	End If
  End If
End Sub
