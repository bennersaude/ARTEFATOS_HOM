'HASH: 9F66BBA7E3709866DDF121C9E9E34D32
' MACRO SAM_ROTINAPORTABILIDADE
'#Uses "*bsShowMessage"

Dim viRetorno As Long
Dim vsMensagemErro As String

Public Sub BOTAOCANCELAR_OnClick()

  If VisibleMode Then
    Dim sp As BStoredProc
    Dim qUpdate As Object
  	Set qUpdate = NewQuery
	Set sp = NewStoredProc

	sp.Name = "BsBen_PortabilidadeCarencia"
	sp.AddParam("P_ACAO", ptInput, ftString)
	sp.AddParam("P_ROTINA", ptInput, ftInteger)
	sp.ParamByName("P_ACAO").AsString = "C"
	sp.ParamByName("P_ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	sp.ExecProc

	qUpdate.Clear
    qUpdate.Add("UPDATE SAM_ROTINAPORTABILIDADE         ")
    qUpdate.Add("   SET SITUACAO = :SITUACAO,           ")
    qUpdate.Add("       USUARIOPROCESSAMENTO = NULL,    ")
    qUpdate.Add("       DATAPROCESSAMENTO = NULL        ")
    qUpdate.Add(" WHERE HANDLE = :HANDLE                ")
    qUpdate.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qUpdate.ParamByName("SITUACAO").AsString = "1"
    qUpdate.ExecSQL

	bsShowMessage("Processo cancelado com sucesso!", "I")

	RefreshNodesWithTable("SAM_ROTINAPORTABILIDADE")

	Set sp = Nothing
	Set qUpdate = Nothing
  Else
  	Dim bsServerExec As Object
	Set bsServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = bsServerExec.ExecucaoImediata(CurrentSystem, _
                                    "SAMPROCEDURE", _
                                    "CancelarPortabilidadeCarencia", _
                                    "Cancelamento dos candidatos à portabilidade de carências", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_ROTINAPORTABILIDADE", _
                                    "SITUACAO", _
                                    "", _
                                    "", _
                                    "C", _
                                    False, _
                                    vsMensagemErro, _
                                    Null)

    If viRetorno = 0 Then
	  bsShowMessage("Processo enviado para execução no servidor!", "I")
	Else
	  bsShowMessage("Erro ao enviar processo para execução§Ã£o no servidor!" + Chr(13) + vsMensagemErro, "I")
	End If

	Set bsServerExec = Nothing
  End If
End Sub

Public Sub BOTAOEMITIRAVISO_OnClick()
	Dim qRelatorio As Object
    Set qRelatorio = NewQuery

    qRelatorio.Clear
    qRelatorio.Add("SELECT HANDLE            ")
	qRelatorio.Add("  FROM R_RELATORIOS      ")
	qRelatorio.Add(" WHERE CODIGO = 'BEN150' ")
	qRelatorio.Active = True
	UserVar("HRotinaBen150") = CurrentQuery.FieldByName("HANDLE").AsString


    ReportPreview(qRelatorio.FieldByName("HANDLE").AsInteger, "", False, False)

    Set qRelatorio = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()

	If VisibleMode Then
	  Dim sp As BStoredProc
	  Dim qUpdate As Object
	  Set qUpdate = NewQuery
	  Set sp = NewStoredProc

	  sp.Name = "BsBen_PortabilidadeCarencia"
	  sp.AddParam("P_ACAO", ptInput, ftString)
	  sp.AddParam("P_ROTINA", ptInput, ftInteger)
	  sp.ParamByName("P_ACAO").AsString = "P"
	  sp.ParamByName("P_ROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	  sp.ExecProc

	  qUpdate.Clear
      qUpdate.Add("UPDATE SAM_ROTINAPORTABILIDADE          ")
      qUpdate.Add("   SET SITUACAO = :SITUACAO,            ")
      qUpdate.Add("       USUARIOPROCESSAMENTO = :USUARIO, ")
      qUpdate.Add("       DATAPROCESSAMENTO = :DATA        ")
      qUpdate.Add(" WHERE HANDLE = :HANDLE                 ")
      qUpdate.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qUpdate.ParamByName("USUARIO").AsInteger = CurrentUser
      qUpdate.ParamByName("DATA").AsDateTime = ServerNow
      qUpdate.ParamByName("SITUACAO").AsString = "5"
      qUpdate.ExecSQL

	  bsShowMessage("Processo concluído com sucesso!", "I")

	  RefreshNodesWithTable("SAM_ROTINAPORTABILIDADE")

	  Set qUpdate = Nothing
	  Set sp = Nothing
	Else
	  Dim bsServerExec As Object
  	  Set bsServerExec = CreateBennerObject("BSServerExec.ProcessosServidor")
	  viRetorno = bsServerExec.ExecucaoImediata(CurrentSystem, _
                                    "SAMPROCEDURE", _
                                    "ProcessarPortabilidadeCarencia", _
                                    "Geração de candidatos à portabilidade de carências", _
                                    CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                    "SAM_ROTINAPORTABILIDADE", _
                                    "SITUACAO", _
                                    "", _
                                    "", _
                                    "P", _
                                    True, _
                                    vsMensagemErro, _
                                    Null)

      If viRetorno = 0 Then
		bsShowMessage("Processo enviado para execução no servidor!", "I")
	  Else
		bsShowMessage("Erro ao enviar processo para execução§Ã£o no servidor!" + Chr(13) + vsMensagemErro, "I")
	  End If

	  Set bsServerExec = Nothing
	End If

End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		UserVar("HRotinaBen150") = CurrentQuery.FieldByName("HANDLE").AsString
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qverifica As Object
    Set qverifica = NewQuery

    qverifica.Clear
    qverifica.Add("SELECT COUNT(1) QTD                  ")
    qverifica.Add("  FROM SAM_ROTINAPORTABILIDADE       ")
    qverifica.Add(" WHERE COMPETENCIABASE = :COMPETENCIA")
    qverifica.ParamByName("COMPETENCIA").AsDateTime = CurrentQuery.FieldByName("COMPETENCIABASE").AsDateTime
    qverifica.Active = True

    If qverifica.FieldByName("QTD").AsInteger > 0 Then
	  bsShowMessage("Já existe uma rotina com a mesma competência inserida!", "E")
      CanContinue = False
      Exit Sub
    End If

    Set qverifica = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	 Select Case CommandID
      Case "BOTAOPROCESSAR"
		BOTAOPROCESSAR_OnClick
      Case "BOTAOCANCELAR"
        BOTAOCANCELAR_OnClick
      Case "BOTAOEMITIRAVISO"
        BOTAOEMITIRAVISO_OnClick
  End Select
End Sub
