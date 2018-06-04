'HASH: 3DF8EE4A968E269B18D408F898C19FFC
'#Uses "*bsShowMessage"

Public Sub PARAMETRIZACAO_OnChange()
	ITEM.LocalWhere = "ANS_SIP_ANEXO_ITEM.SIPANEXO = " + CurrentQuery.FieldByName("PARAMETRIZACAO").AsString
End Sub

Public Sub PARAMETRIZACAO_OnExit()
	If CurrentQuery.FieldByName("ROTINASIP").IsNull Then
		ITEM.LocalWhere = ""
    End If
End Sub

Public Sub ROTINASIP_OnChange()
	PARAMETRIZACAO.LocalWhere = "ANS_SIP_ANEXO.RESOLUCAO = " + CurrentQuery.FieldByName("ROTINASIP").AsString
End Sub

Public Sub ROTINASIP_OnExit()
	If CurrentQuery.FieldByName("ROTINASIP").IsNull Then
		PARAMETRIZACAO.LocalWhere = ""
		ITEM.LocalWhere = ""
    End If
End Sub

Public Sub TABLE_AfterPost()
  Dim vsMensagemRetorno As String
  If VisibleMode Then
    	Dim interface As Object

    	Set interface = CreateBennerObject("BSINTERFACE0055.ImportarParametros")
	    interface.ImportarParametros(CurrentSystem, _
	    							 RecordHandleOfTable("ANS_SIP_ANEXO_ITEM"), _
	                   				 CurrentQuery.FieldByName("ROTINASIP").AsInteger, _
	                   				 CurrentQuery.FieldByName("PARAMETRIZACAO").AsInteger, _
	                   				 CurrentQuery.FieldByName("ITEM").AsInteger, _
	                   				 vsMensagemRetorno)
    	Set interface = Nothing

  ElseIf WebMode Then

 		Dim vsMensagemErro As String
    	Dim viRetorno As Long
    	Dim Obj As Object

        Dim vcContainer As CSDContainer
   		Set vcContainer = NewContainer
   		vcContainer.AddFields("HANDLE:INTEGER;ROTINASIP:INTEGER;" + _
   							  "PARAMETRIZACAO:INTEGER;ITEM:INTEGER")

   		vcContainer.Insert
	   	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO_ITEM")
   		vcContainer.Field("ROTINASIP").AsInteger = CurrentQuery.FieldByName("ROTINASIP").AsInteger
   		vcContainer.Field("PARAMETRIZACAO").AsInteger = CurrentQuery.FieldByName("PARAMETRIZACAO").AsInteger
   		vcContainer.Field("ITEM").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger

	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     	"BSANS001", _
                                     	"ImportarParametros", _
                                     	"SIP - Sistema de informações de Produtos - Importar Parâmetros", _
                                        RecordHandleOfTable("ANS_SIP_ANEXO_ITEM"), _
                                     	"ANS_SIP_ANEXO_ITEM", _
                                     	"SITUACAOPROCESSO", _
                                     	"", _
                                     	"", _
                                     	"P", _
                                     	True, _
	                                   	vsMensagemErro, _
                                     	vcContainer, _
                                     	False)

	    If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If

    	Set Obj = Nothing
  End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

 If CurrentQuery.FieldByName("ROTINASIP").IsNull Then
    bsShowMessage("Selecione a Rotina SIP.", "E")
    CanContinue = False
    Exit Sub
 End If

 If CurrentQuery.FieldByName("PARAMETRIZACAO").IsNull Then
    bsShowMessage("Selecione a Parametrização.", "E")
    CanContinue = False
    Exit Sub
 End If

 If CurrentQuery.FieldByName("ROTINASIP").IsNull Then
    bsShowMessage("Selecione o Item.", "E")
    CanContinue = False
    Exit Sub
 End If

End Sub

