'HASH: 58243A4A88149AEC0D2F0565FB62C168
'#Uses "*PrimeiroDiaCompetencia"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qVerificaDataProcessamento As BPesquisa
	Set qVerificaDataProcessamento = NewQuery
	qVerificaDataProcessamento.Active = False
	qVerificaDataProcessamento.Clear
	qVerificaDataProcessamento.Add("SELECT PROVISORIO, COMPETENCIA ")
	qVerificaDataProcessamento.Add("  FROM FIS_REGAUXCOMPETENCIA   ")
	qVerificaDataProcessamento.Add(" WHERE HANDLE = :HANDLE        ")
	qVerificaDataProcessamento.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("FIS_REGAUXCOMPETENCIA")
	qVerificaDataProcessamento.Active = True

	If qVerificaDataProcessamento.FieldByName("PROVISORIO").AsString = "N" Then
		Dim vQuery As BPesquisa
		Set vQuery = NewQuery
		vQuery.Clear
		vQuery.Add("SELECT PERIODOFATCONINICIAL FROM SFN_PARAMETROSFIN")
		vQuery.Active = True

		If (PrimeiroDiaCompetencia(qVerificaDataProcessamento.FieldByName("COMPETENCIA").AsDateTime) >= PrimeiroDiaCompetencia(vQuery.FieldByName("PERIODOFATCONINICIAL").AsDateTime)) Then
			bsShowMessage("Processar registro auxiliar com competência dentro do período contábil só é permitido quando registro provisório.", "E")
			CanContinue = False
			Set qVerificaDataProcessamento = Nothing
			Set vQuery = Nothing
			Exit Sub
		End If
		Set vQuery = Nothing
	End If
	Set qVerificaDataProcessamento = Nothing

  	If VisibleMode Then
		Dim interface As Object
		Set interface = CreateBennerObject("BSINTERFACE0059.RegistrosAuxiliares_Processamento")
		interface.Inicializar(CurrentSystem)

  		Dim vHandleFisComp As Long
  		vHandleFisComp = RecordHandleOfTable("FIS_REGAUXCOMPETENCIA")
		
	    interface.Exec(CurrentSystem, vHandleFisComp, _
	                   CurrentQuery.FieldByName("CHECK01").AsString, _
	                   CurrentQuery.FieldByName("CHECK02").AsString, _
	                   CurrentQuery.FieldByName("CHECK03").AsString, _
	                   CurrentQuery.FieldByName("CHECK04").AsString, _
	                   CurrentQuery.FieldByName("CHECK05").AsString, _
	                   CurrentQuery.FieldByName("CHECK06").AsString)
		interface.Finalizar
	    Set interface = Nothing
	ElseIf WebMode Then
		Dim vsMensagemErro As String
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

		Dim vDll As Object
		Set vDll = CreateBennerObject("BSServerExec.ProcessosServidor")

		Dim viRetorno As Long
		viRetorno = vDll.ExecucaoImediata(CurrentSystem, _
                                    	"FIS001", _
                                     	"RegistrosAuxiliares_Processamento", _
                                     	"Processamento dos Registros Auxiliares", _
                                        RecordHandleOfTable("FIS_REGAUXCOMPETENCIA"), _
                                     	"FIS_REGAUXCOMPETENCIA", _
                                     	"SITUACAOPROCESSO", _
                                     	"", _
                                     	"", _
                                     	"P", _
                                     	True, _
	                                   	vsMensagemErro, _
                                     	vcContainer, _
                                     	False)
		Set vDll = Nothing
		If viRetorno = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
		Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		End If
	End If
End Sub
