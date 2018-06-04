'HASH: 1FA49146994B98D4BA4FF916017AC26D
 
'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then

   		Dim vsMensagemErro As String
   		Dim viRetorno As Long
   		Dim vcContainer As CSDContainer
   		Set vcContainer = NewContainer
   		vcContainer.AddFields("HANDLE:INTEGER;ARQUIVO:INTEGER;EMAIL:STRING;NOMEARQUIVO:STRING")


   		vcContainer.Insert
	   	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINARESUMO")
   		vcContainer.Field("ARQUIVO").AsInteger = CurrentQuery.FieldByName("RADIOARQUIVO").AsInteger
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
                                    	"RotinaResumo_ExportacaoRec", _
                                    	"Exportação do resumo de IRRF de recebimento - " +	SQL.FieldByName("DESCRICAO").AsString, _
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
