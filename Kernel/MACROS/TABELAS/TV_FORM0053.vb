'HASH: F7BE2ADEFB04E866C4856D631B810AF9

'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"

Dim viHGuia As Long
Dim viHPeg As Long

Public Sub TABLE_AfterPost()
  Dim vvSamPegDigit As Object
  Dim vsMsg As String
  Dim vbCriouGlosa As Boolean
  Dim Obj As Object
  Dim viRetorno As Long
  Dim vsMensagemErro As String
  Dim vvContainer As CSDContainer

  CriaTabelaTemporariaSqlServer

  If (WebMode) Then

  	Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Add("SELECT PEG, GUIAAUXILIAR")
    qSQL.Add("FROM SAM_GUIA")
    qSQL.Add("WHERE HANDLE = :HGUIA")
    qSQL.ParamByName("HGUIA").AsInteger = viHGuia
    qSQL.Active = True

    Set vvContainer = NewContainer

    vvContainer.AddFields("GUIA:INTEGER;")
	vvContainer.AddFields("PEG:INTEGER;")
    vvContainer.AddFields("MOTIVOGLOSA:INTEGER;")
    vvContainer.AddFields("COMPLEMENTO:STRING;")

    vvContainer.Insert
    vvContainer.Field("GUIA").AsInteger = viHGuia
	vvContainer.Field("PEG").AsInteger = 0
    vvContainer.Field("MOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    vvContainer.Field("COMPLEMENTO").AsString = CurrentQuery.FieldByName("COMPLEMENTO").AsString

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                            		"SAMPEGDIGIT", _
	                                     "Rotinas_GlosaTotal", _
	                                     "Glosa Total da Guia " + qSQL.FieldByName("GUIAAUXILIAR").AsString + _
                                     	 " | PEG n. " + CStr(qSQL.FieldByName("PEG").AsInteger), _
	                                     0, _
	                                     "", _
	                                     "", _
	                                     "", _
	                                     "", _
	                                     "", _
	                                     True, _
	                                     vsMensagemErro, _
	                                     vvContainer)
    If viRetorno = 0 Then
	 	bsShowMessage("Processo enviado para execução no servidor!", "I")
	Else
	   	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	End If

	Set Obj = Nothing

  Else
	Set vvSamPegDigit = CreateBennerObject("SAMPEGDIGIT.Rotinas_GlosaTotal")

    vsMsg = vvSamPegDigit.GlosaTotal(CurrentSystem, _
									   viHGuia, _
									   CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger, _
									   CurrentQuery.FieldByName("COMPLEMENTO").AsString, _
  									   vbCriouGlosa)
  	If vsMsg <> "" Then
	  bsShowMessage(vsMsg, "I")
    End If

    Set vvSamPegDigit = Nothing
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  viHGuia = RecordHandleOfTable("SAM_GUIA")

  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Add("SELECT SITUACAO,PEG")
  qSQL.Add("FROM SAM_GUIA")
  qSQL.Add("WHERE HANDLE = :HGUIA")
  qSQL.ParamByName("HGUIA").AsInteger = viHGuia
  qSQL.Active = True

  viHPeg = qSQL.FieldByName("PEG").AsInteger

  Dim agrupadorFechado As Boolean

  agrupadorFechado = VerificaAgrupadorPagamentoFechado(viHPeg)

  If (agrupadorFechado) Then
  	BsShowMessage("Não é permitida a alteração de glosa ligada à registro de pagamento fechado.","E")
    CanContinue = False
    Exit Sub
  End If

  'Em modo desktop este tratamento será realizado no evento do click do botão BOTAOGLOSATOTAL da SAM_GUIA
  If WebMode Then
    'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
    If qSQL.FieldByName("SITUACAO").AsString = "1" Then 'digitacao
      If Not CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "A") Then
        bsShowMessage("Usuário não tem permissão para alterar nesta filial!", "E")
        CanContinue = False
      End If
    Else
      If Not CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "P") Then
        bsShowMessage("Usuário não tem permissão nesta filial!", "E")
      End If
    End If
    Set qSQL = Nothing
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Trim(CurrentQuery.FieldByName("COMPLEMENTO").AsString) = "" Then
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Add("SELECT EXIGECOMPLEMENTO")
    qSQL.Add("FROM SAM_MOTIVOGLOSA")
    qSQL.Add("WHERE HANDLE = :HMOTIVOGLOSA")
    qSQL.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("EXIGECOMPLEMENTO").AsString = "S" Then
		bsShowMessage("Motivo de glosa exige complemento", "E")
		CanContinue = False
    End If
    Set qSQL = Nothing
  End If
End Sub

Public Function VerificaAgrupadorPagamentoFechado(pPeg As Long) As Boolean
	Dim callEntity As CSEntityCall
  	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamPeg, Benner.Saude.Entidades", "VerificaPegVinculadoPagamentoFechado")
  	callEntity.AddParameter(pdtAutomatic, pPeg)
  	VerificaAgrupadorPagamentoFechado = CBool(callEntity.Execute)
	Set callEntity =  Nothing
End Function
