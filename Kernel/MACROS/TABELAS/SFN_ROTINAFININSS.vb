'HASH: 2EACE78A174852F1FCDF3E5F88A7C4AF

'Macro: SFN_ROTINAFINFAT_INSS
'#Uses "*bsShowMessage"
'#Uses "*ProcuraPrestador"

Dim vsMensagem As String


Public Sub BOTAOCANCELAR_OnClick()
  If bsShowMessage("Confirma o cancelamento da rotina ?", "Q") = vbYes Then
    If CurrentQuery.State <>1 Then
      bsShowMessage("Os parâmetros não podem estar em edição", "I")
      Exit Sub
    End If

    Dim Obj As Object

    'Rafael Zarpellon - SMS 90431 - 03/04/2008 - Início
    'Set INSS = CreateBennerObject("SAMINSS.INSS")
    'INSS.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    'Set INSS = Nothing
    If VisibleMode Then
      Set Obj = CreateBennerObject("BSINTERFACE0038.RotinaINSS")
      Obj.Cancela(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Else
      Dim viRetorno As Long
      Dim qSQL As Object
      Set qSQL = NewQuery

      qSQL.Clear
      qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
      qSQL.Add("       CFIN.COMPETENCIA,")
      qSQL.Add("       RFIN.SEQUENCIA")
      qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
      qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
      qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
      qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
      qSQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
      qSQL.Active = True

      Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                       "SAMINSS", _
                                       "RotinaINSS_CancelaRotina", _
                                       "Rotina de Faturamento de INSS (Cancelamento) -" + _
                                         " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                         " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                       CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                       "SFN_ROTINAFININSS", _
                                       "SITUACAO", _
                                       "", _
                                       "", _
                                       "C", _
                                       False, _
                                       vsMensagem, _
                                       Null)

      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If

      Set qSQL = Nothing
    End If

    Set Obj = Nothing
    'Rafael Zarpellon - SMS 90431 - 03/04/2008 - Fim
  End If
End Sub

Public Sub BOTAOGERARSEFIP_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim Obj As Object

  'Rafael Zarpellon - SMS 90431 - 04/04/2008 - Início
  'Set INSS = CreateBennerObject("SAMINSS.INSS")
  'INSS.GerarSEFIP(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  'Set INSS = Nothing
  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0038.RotinaINSS")
    Obj.GerarSEFIP(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim viRetorno As Long
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Clear
    qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    qSQL.Add("       CFIN.COMPETENCIA,")
    qSQL.Add("       RFIN.SEQUENCIA")
    qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
    qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    qSQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qSQL.Active = True

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SAMINSS", _
                                     "RotinaINSS_GerarSEFIPRotina", _
                                     "Rotina de Faturamento de INSS (Arquivo SEFIP) -" + _
                                       " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                       " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFININSS", _
                                     "SITUACAOARQUIVOSEFIP", _
                                     "SITUACAO", _
                                     "Rotina ainda não foi processada!", _
                                     "P", _
                                     True, _
                                     vsMensagem, _
                                     Null)

      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If

      Set qSQL = Nothing
   End If
   Set Obj = Nothing
  'Rafael Zarpellon - SMS 90431 - 01/04/2008 - Início
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  Dim Obj As Object

  'Rafael Zarpellon - SMS 90431 - 01/04/2008 - Início
  'Set INSS = CreateBennerObject("SAMINSS.INSS")
  'INSS.Faturar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  'Set INSS = Nothing
  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0038.RotinaINSS")
    Obj.Processa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim viRetorno As Long
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Clear
    qSQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    qSQL.Add("       CFIN.COMPETENCIA,")
    qSQL.Add("       RFIN.SEQUENCIA")
    qSQL.Add("FROM SFN_ROTINAFIN       RFIN")
    qSQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    qSQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    qSQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    qSQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qSQL.Active = True

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SAMINSS", _
                                     "RotinaINSS_Faturar", _
                                     "Rotina de Faturamento de INSS (Processamento) -" + _
                                       " Competência: " + Str(Format(qSQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                       " Sequência: "   + qSQL.FieldByName("SEQUENCIA").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFININSS", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagem, _
                                     Null)

      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If

      Set qSQL = Nothing
   End If
   Set Obj = Nothing
   'Rafael Zarpellon - SMS 90431 - 01/04/2008 - Fim
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim viHandlePrestador As Long
  Dim vsCPFNome As String

  ShowPopup = False

  If (IsNumeric(PRESTADOR.LocateText)) Then
    vsCPFNome = "C"
  Else
    vsCPFNome = "N"
  End If

  viHandlePrestador = ProcuraPrestador(vsCPFNome, "T", PRESTADOR.LocateText)

  If viHandlePrestador <> 0 Then
       CurrentQuery.Edit
       CurrentQuery.FieldByName("PRESTADOR").Value = viHandlePrestador
  End If

End Sub

Public Sub TABLE_AfterScroll()
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Início
  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery
    qRotinaFin.Active = False
    qRotinaFin.Clear
    qRotinaFin.Add("SELECT SITUACAO")
    qRotinaFin.Add("  FROM SFN_ROTINAFIN")
    qRotinaFin.Add(" WHERE HANDLE = :HANDLE")
    qRotinaFin.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    qRotinaFin.Active = True
  If (qRotinaFin.FieldByName("SITUACAO").AsString = "P") Or (qRotinaFin.FieldByName("SITUACAO").AsString = "S") Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If
  Set qRotinaFin = Nothing
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Fim
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("Não foi possível excluir, a rotina está processada !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
    bsShowMessage("Alteração negada, a rotina não está aberta !", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERARSEFIP"
			BOTAOGERARSEFIP_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
