'HASH: D6766E70DA7BB831116347D0C477893F
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()


  If VisibleMode Then
    	Dim interface As Object

    	Set interface = CreateBennerObject("BSINTERFACE0059.RegistrosAuxiliares_Cancelamento")
	    interface.Exec(CurrentSystem, RecordHandleOfTable("FIS_REGAUXCOMPETENCIA"), _
	                   CurrentQuery.FieldByName("CHECK01").AsString, _
	                   CurrentQuery.FieldByName("CHECK02").AsString, _
	                   CurrentQuery.FieldByName("CHECK03").AsString, _
	                   CurrentQuery.FieldByName("CHECK04").AsString, _
	                   CurrentQuery.FieldByName("CHECK05").AsString, _
	                   CurrentQuery.FieldByName("CHECK06").AsString)
    	Set interface = Nothing

  ElseIf WebMode Then

 		Dim vsMensagemErro As String
    	Dim viRetorno As Long

        Dim vcContainer As CSDContainer
   		Set vcContainer = NewContainer
   		vcContainer.AddFields("HANDLE:INTEGER;CHECK01:STRING;" + _
   							  "CHECK02:STRING;CHECK03:STRING;CHECK04:STRING;CHECK05:STRING;CHECK06:STRING")

   		vcContainer.Insert
	   	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("FIS_REGAUXCOMPETENCIA")
   		vcContainer.Field("CHECK01").AsString = CurrentQuery.FieldByName("CHECK01").AsString
   		vcContainer.Field("CHECK02").AsString = CurrentQuery.FieldByName("CHECK02").AsString
   		vcContainer.Field("CHECK03").AsString = CurrentQuery.FieldByName("CHECK03").AsString
   		vcContainer.Field("CHECK04").AsString = CurrentQuery.FieldByName("CHECK04").AsString
   		vcContainer.Field("CHECK05").AsString = CurrentQuery.FieldByName("CHECK05").AsString
   		vcContainer.Field("CHECK06").AsString = CurrentQuery.FieldByName("CHECK06").AsString

	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     	"FIS001", _
                                     	"RegistrosAuxiliares_Cancelamento", _
                                     	"Cancelamento dos Registros Auxiliares", _
                                        RecordHandleOfTable("FIS_REGAUXCOMPETENCIA"), _
                                     	"FIS_REGAUXCOMPETENCIA", _
                                     	"SITUACAOPROCESSO", _
                                     	"", _
                                     	"", _
                                     	"C", _
                                     	True, _
	                                   	vsMensagemErro, _
                                     	vcContainer, _
                                     	False)

	    If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If
  End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificaDataProcessamento As Object
  Set qVerificaDataProcessamento = NewQuery

  qVerificaDataProcessamento.Active = False
  qVerificaDataProcessamento.Clear

  qVerificaDataProcessamento.Add("SELECT PROVISORIO, DATAPROCESSAMENTOROTINA, DATAPROCESSAMENTO ")
  qVerificaDataProcessamento.Add("  FROM FIS_REGAUXCOMPETENCIA                                  ")
  qVerificaDataProcessamento.Add(" WHERE HANDLE = :HANDLE                                       ")

  qVerificaDataProcessamento.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("FIS_REGAUXCOMPETENCIA")

  qVerificaDataProcessamento.Active = True

  If qVerificaDataProcessamento.FieldByName("DATAPROCESSAMENTOROTINA").IsNull And qVerificaDataProcessamento.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("Rotina já se encontra em aberto! Impossível cancelar.", "E")
    CanContinue = False
    Set qVerificaDataProcessamento = Nothing
    Exit Sub
  End If

  If qVerificaDataProcessamento.FieldByName("PROVISORIO").AsString = "N" Then
    Set Query = NewQuery

    Query.Clear
    Query.Add("SELECT DIASLIBERACAOCANCELREGAUX FROM SFN_PARAMETROSFIN")
    Query.Active = True

    If (Query.FieldByName("DIASLIBERACAOCANCELREGAUX").AsInteger < (ServerDate - qVerificaDataProcessamento.FieldByName("DATAPROCESSAMENTO").AsDateTime)) Then
    	bsShowMessage("O prazo máximo permitido para o cancelamento expirou.", "E")
        CanContinue = False
    	Set Query = Nothing
    	Set qVerificaDataProcessamento = Nothing
    	Exit Sub

  	End If

    Set Query = Nothing

  End If

  Set qVerificaDataProcessamento = Nothing



End Sub
