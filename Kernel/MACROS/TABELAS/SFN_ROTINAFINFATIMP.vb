'HASH: 9084464056840694A78A436988BFB9A4

'macro sfn_rotinafinfatimp - fernando sms 30330 - 04/09/2004
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()

  'PARA VERIFICAR SE JAH FOI PROCESSADA
  Dim qVerificarSituacao As Object
  Dim Obj As Object
  Set qVerificarSituacao = NewQuery

  qVerificarSituacao.Active = False
  qVerificarSituacao.Clear
  qVerificarSituacao.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  qVerificarSituacao.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  qVerificarSituacao.Active = True
  If qVerificarSituacao.FieldByName("SITUACAO").Value = "P" Then
    qVerificarSituacao.Active = False
    Set qVerificarSituacao = Nothing
    bsShowMessage("A Rotina já foi processada", "I")
    Exit Sub
  End If

  Set qVerificarSituacao = Nothing

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0069.FaturamentoImportacao")
    Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "VISUAL")
  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long

    Dim vcContainer As CSDContainer
    Set vcContainer = NewContainer
    vcContainer.AddFields("HANDLE:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen012", _
                                     "FaturamentoImportacao_Processar", _
                                     "Processando Rotina de Faturamento por Importação", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINFATIMP", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     vcContainer)

    Set SQL = Nothing
    Set vcContainer = Nothing

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set Obj = Nothing
End Sub

Public Sub BOTATOCANCELAR_OnClick()
  Dim qVerificarSituacao As Object
  Dim Obj As Object
  Set qVerificarSituacao = NewQuery

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
    Exit Sub
  End If

  qVerificarSituacao.Active = False
  qVerificarSituacao.Clear
  qVerificarSituacao.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  qVerificarSituacao.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  qVerificarSituacao.Active = True
  If qVerificarSituacao.FieldByName("SITUACAO").Value = "A" Then
    qVerificarSituacao.Active = False
    Set qVerificarSituacao = Nothing
    bsShowMessage("A Rotina não está processada", "I")
    Exit Sub
  End If
  'qVerificarSituacao.Active = False
  Set qVerificarSituacao = Nothing

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0069.FaturamentoImportacao")
    Obj.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long

    Dim vcContainer As CSDContainer
    Set vcContainer = NewContainer
    vcContainer.AddFields("HANDLE:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen012", _
                                     "FaturamentoImportacao_Cancelar", _
                                     "Cancelando Rotina de Faturamento por Importação", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINFATIMP", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                     vsMensagemErro, _
                                     vcContainer)

    Set SQL = Nothing
    Set vcContainer = Nothing

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set Obj = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub


Public Sub VerificaSeProcessada(CanContinue As Boolean)
  Dim qVerificarSituacao As Object
  Set qVerificarSituacao = NewQuery

  qVerificarSituacao.Active = False
  qVerificarSituacao.Clear
  qVerificarSituacao.Add("SELECT SITUACAO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE")
  qVerificarSituacao.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  qVerificarSituacao.Active = True
  If qVerificarSituacao.FieldByName("SITUACAO").Value = "P" Then
    CanContinue = False
    qVerificarSituacao.Active = False
    Set qVerificarSituacao = Nothing
    bsShowMessage("A Rotina já foi processada", "E")
    Exit Sub
  End If

  'qVerificarSituacao.Active = False
  Set qVerificarSituacao = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
		Case "BOTATOCANCELAR"
			BOTATOCANCELAR_OnClick
	End Select
End Sub
