'HASH: 6A0425753A147D9C7171C0F23B6E665A
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("TESOURARIA").AsInteger = RecordHandleOfTable("SFN_TESOURARIA")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If CurrentQuery.FieldByName("TESOURARIA").AsInteger = CurrentQuery.FieldByName("TRANSFTESOURARIA").AsInteger Then
		bsShowMessage("Tesouraria ORIGEM não pode a mesma tesouraria de DESTINO!", "E")
		CanContinue = False
		Exit Sub
	End If



    Dim Interface As Object
	Dim vsMensagem As String


    Set Interface = CreateBennerObject("SFNTesouraria.Tesouraria")
	vsMensagem = Interface.Transferencia(CurrentSystem, CurrentQuery.FieldByName("DATA").AsDateTime, CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
    						CurrentQuery.FieldByName("TESOURARIA").AsInteger, CurrentQuery.FieldByName("TRANSFTESOURARIA").AsInteger, _
    						CurrentQuery.FieldByName("VALOR").AsFloat,CurrentQuery.FieldByName("CHEQUE").AsInteger,CurrentQuery.FieldByName("HISTORICO").AsString)
    Set Interface = Nothing

	If vsMensagem <> "" Then
		bsShowMessage(vsMensagem, "I")
	End If

End Sub
