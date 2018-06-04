'HASH: C0200EB663570E368C3616C2025177D8
'#Uses "*bsShowMessage"


Public Sub TABLE_AfterScroll()

	Dim SQL As Object

	Set SQL = NewQuery
	SQL.Add("SELECT NUMERORECIBO FROM SFN_ROTINARESUMO WHERE HANDLE = :HANDLE")
	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
	SQL.Active = True


	If SQL.FieldByName("NUMERORECIBO").IsNull Then
		CurrentQuery.FieldByName("CHECKDECLARACAORETIFICADORA").AsString = "N"
	Else
		CurrentQuery.FieldByName("CHECKDECLARACAORETIFICADORA").AsString = "S"

	End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then

   		Dim vsMensagemErro As String
   		Dim viRetorno As Long
   		Dim vcContainer As CSDContainer
   		Set vcContainer = NewContainer
   		vcContainer.AddFields("HANDLE:INTEGER;CNPJDECLARANTE:STRING;NOMEEMPDECLARANTE:STRING;RADIODECLARANTE:INTEGER;CHECKDECLARACAO:BOOLEAN;" + _
   							  "CPFRESPONSAVELDECLARANTE:STRING; ANOREFERENCIA:STRING; NUMRECIBODECLARACAO:STRING;" + _
   							  "CPF:STRING;NOME:STRING;DDD:STRING; TELEFONE:STRING;RAMAL:STRING; FAX:STRING;EMAIL:STRING; NOMEARQUIVO:STRING")


   		vcContainer.Insert
	   	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
   		vcContainer.Field("CNPJDECLARANTE").AsString = CurrentQuery.FieldByName("CNPJDECLARANTE").AsString
   		vcContainer.Field("NOMEEMPDECLARANTE").AsString = CurrentQuery.FieldByName("NOMEEMPDECLARANTE").AsString
   		vcContainer.Field("RADIODECLARANTE").AsInteger = CurrentQuery.FieldByName("RADIODECLARANTE").AsInteger
   		vcContainer.Field("CPFRESPONSAVELDECLARANTE").AsString = CurrentQuery.FieldByName("CPFRESPONSAVELDECLARANTE").AsString
   		vcContainer.Field("ANOREFERENCIA").AsString = CurrentQuery.FieldByName("ANOREFERENCIA").AsString
   		vcContainer.Field("CPF").AsString = CurrentQuery.FieldByName("CPF").AsString
   		vcContainer.Field("NOME").AsString = CurrentQuery.FieldByName("NOME").AsString
   		vcContainer.Field("DDD").AsString = CurrentQuery.FieldByName("DDD").AsString
   		vcContainer.Field("TELEFONE").AsString = CurrentQuery.FieldByName("TELEFONE").AsString
   		vcContainer.Field("RAMAL").AsString = CurrentQuery.FieldByName("RAMAL").AsString
   		vcContainer.Field("FAX").AsString = CurrentQuery.FieldByName("FAX").AsString
   		vcContainer.Field("EMAIL").AsString = CurrentQuery.FieldByName("EMAIL").AsString
   		vcContainer.Field("NOMEARQUIVO").AsString = CurrentQuery.FieldByName("NOMEARQUIVO").AsString


		Dim SQL As Object
		Set SQL = NewQuery
		SQL.Add("SELECT DESCRICAO FROM SFN_ROTINARESUMO WHERE HANDLE = :HANDLE")
		SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
		SQL.Active = True

   		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
   		viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                    	"BSFIN006", _
                                    	"RotinaResumo_ExportacaoDIRF", _
                                    	"Geração de Arquivo DIRF - " +	SQL.FieldByName("DESCRICAO").AsString, _
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

