'HASH: B6AC729AD2C24BD943DA0C2B4FB7FF3F


'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then

   		Dim vsMensagemErro As String
   		Dim viRetorno As Long
   		Dim vcContainer As CSDContainer
   		Set vcContainer = NewContainer
   		vcContainer.AddFields("HANDLE:INTEGER;CONTFININICIAL:INTEGER;CONTFINFINAL:INTEGER;CONTASJAPROCESSADAS:STRING;GRUPOCONTRATO:STRING;CONTRATO:STRING")

		Dim vsGrupoContrato As String
		Dim vsContrato As String

		Dim SQL As Object
		Set SQL = NewQuery
		SQL.Add("SELECT DESCRICAO, GRUPOCONTRATO, CONTRATO FROM SFN_ROTINARESUMO WHERE HANDLE = :HANDLE")
		SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
		SQL.Active = True

		vsGrupoContrato = SQL.FieldByName("GRUPOCONTRATO").AsString
		vsGrupoContrato =  Replace(Replace( vsGrupoContrato, "|_|", "," ), "|_", "")

		vsContrato = SQL.FieldByName("CONTRATO").AsString
		vsContrato = Replace(Replace( vsContrato, "|_|", "," ), "|_", "")

   		vcContainer.Insert
	   	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
   		vcContainer.Field("CONTFININICIAL").AsInteger = CurrentQuery.FieldByName("CONTAFININICIAL").AsInteger
   		vcContainer.Field("CONTFINFINAL").AsInteger = CurrentQuery.FieldByName("CONTAFINFINAL").AsInteger
   		vcContainer.Field("CONTASJAPROCESSADAS").AsString = CurrentQuery.FieldByName("CHECKCONTASJAPROCESSADAS").AsString
   		vcContainer.Field("GRUPOCONTRATO").AsString = vsGrupoContrato
   		vcContainer.Field("CONTRATO").AsString = vsContrato


   		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
   		viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    	"BSFIN006", _
                                    	"RotinaResumo_ProcessarRec", _
                                    	"Processando resumo de IRRF recebimento - " +	SQL.FieldByName("DESCRICAO").AsString, _
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
