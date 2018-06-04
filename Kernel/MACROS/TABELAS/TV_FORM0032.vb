'HASH: 4516ECE23E761E1A8AD55B5A5E006E41
 

'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then

   		Dim vsMensagemErro As String
   		Dim viRetorno As Long
   		Dim vcContainer As CSDContainer
   		Set vcContainer = NewContainer
   		vcContainer.AddFields("HANDLE:INTEGER;CONTFININICIAL:INTEGER;CONTFINFINAL:INTEGER;VERIFICASMOVIMENTACAO:STRING")


   		vcContainer.Insert
	   	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
   		vcContainer.Field("CONTFININICIAL").AsInteger = CurrentQuery.FieldByName("CONTFININICIAL").AsInteger
   		vcContainer.Field("CONTFINFINAL").AsInteger = CurrentQuery.FieldByName("CONTFINFINAL").AsInteger
   		vcContainer.Field("VERIFICASMOVIMENTACAO").AsString = CurrentQuery.FieldByName("CHECKVERIFICASEMMOVIMENTACAO").AsString


		Dim SQL As Object
		Set SQL = NewQuery
		SQL.Add("SELECT DESCRICAO FROM SFN_ROTINARESUMO WHERE HANDLE = :HANDLE")
		SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
		SQL.Active = True


   		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
   		viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    	"BSFIN006", _
                                    	"RotinaResumo_ProcessarPag", _
                                    	"Processando resumo de IRRF pagamento - " +	SQL.FieldByName("DESCRICAO").AsString, _
                                    	0, _
                                    	"SFN_ROTINARESUMO", _
                                    	"", _
                                    	"", _
                                    	"", _
                                    	"P", _
                                    	False, _
                                    	vsMensagemErro, _
                                    	vcContainer)

   		If viRetorno = 0 Then
    	 	bsShowMessage("Processo enviado para execução no servidor!", "I")
   		Else
	     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	   	End If



		Set Obj = Nothing
		Set SQL = Nothing


	End If
End Sub
