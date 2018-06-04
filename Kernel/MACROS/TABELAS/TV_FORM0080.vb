'HASH: 39E3F8F7FA3A3E97AAF0283E2C63B160
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
	Dim vsMensagemErro As String
	Dim viRetorno As Long
	Dim Obj As Object

	Dim vcContainer As CSDContainer
	Set vcContainer = NewContainer
	vcContainer.AddFields("HANDLE:INTEGER;HANDLEDESTINO:INTEGER;" + _
						  "CHECKEVENTO:STRING;CHECKCLASSE:STRING;CHECKCAMPOSIP:STRING;CHECKAPROPRIACAONULA:STRING")

	vcContainer.Insert
	vcContainer.Field("HANDLE").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO")
	vcContainer.Field("HANDLEDESTINO").AsInteger = CurrentQuery.FieldByName("ANEXODESTINO").AsInteger
	vcContainer.Field("CHECKEVENTO").AsString = CurrentQuery.FieldByName("CHECKEVENTO").AsString
	vcContainer.Field("CHECKCLASSE").AsString = CurrentQuery.FieldByName("CHECKCLASSE").AsString
	vcContainer.Field("CHECKCAMPOSIP").AsString = CurrentQuery.FieldByName("CHECKCAMPOSIP").AsString
	vcContainer.Field("CHECKAPROPRIACAONULA").AsString = CurrentQuery.FieldByName("CHECKAPROPRIACAONULA").AsString


	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                             	"BSANS001", _
	                             	"DuplicarAnexo", _
	                             	"SIP - Sistema de informações de Produtos - Duplicação de anexo", _
	                                RecordHandleOfTable("ANS_SIP_ANEXO"), _
	                             	"ANS_SIP_ANEXO", _
	                             	"SITUACAO", _
	                             	"", _
	                             	"", _
	                             	"P", _
	                             	True, _
	                               	vsMensagemErro, _
	                             	vcContainer, _
	                             	False)
	Set Obj = Nothing
	If viRetorno = 0 Then
		bsShowMessage("Processo enviado para execução no servidor!", "I")
	Else
		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("ANEXODESTINO").IsNull Then
		bsShowMessage("O anexo de destino deve ser preenchido!","E")
		CancContinue = False
		Exit Sub
	End If

	If bsShowMessage("A duplicação de modelos apagará todos os dados do modelo destino, Continuar o processo?", "Q") = vbNo Then
		bsShowMessage("Duplicação cancelada","E")
		CancContinue = False
		Exit Sub
	End If


	Dim qSelect As Object
	Set qSelect = NewQuery

	qSelect.Active = False
  	qSelect.Clear

	qSelect.Add("SELECT A.HANDLE            ")
  	qSelect.Add("  FROM ANS_SIP_ANEXO A,    ")
  	qSelect.Add("       ANS_SIP_ANEXO B     ")
  	qSelect.Add(" WHERE A.HANDLE = :HORIGEM ")
  	qSelect.Add(" AND B. HANDLE = :HDESTINO ")
  	qSelect.Add(" AND A.ANEXO  = B.ANEXO    ")

  	qSelect.ParamByName("HORIGEM").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO")
  	qSelect.ParamByName("HDESTINO").AsInteger = CurrentQuery.FieldByName("ANEXODESTINO").AsInteger

  	qSelect.Active = True

  	If qSelect.EOF Then
  		bsShowMessage("O modelo de origem e destino são de anexos diferentes. Duplicação cancelada!","E")
  		CanContinue = False
		Set qSelect = Nothing
  		Exit Sub
	End If

	Set qSelect = Nothing

End Sub
