'HASH: 7E5A30DD27B027038BCFBFCC9799AD58
 '#Uses "*bsShowMessage"
 Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qInsert As Object
	Dim i As Long
	Dim CIDsArray() As String
	Dim vsCampos As String
	Dim vsParametros As String

	If Not CurrentQuery.FieldByName("IDADEINICIAL").IsNull Then
		If Not CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
			If CurrentQuery.FieldByName("IDADEINICIAL").AsInteger > CurrentQuery.FieldByName("IDADEFINAL").AsInteger Then
				bsShowMessage("Idade Final deve ser maior ou igual à Idade Inicial","I")
				Exit Sub
			End If
		Else
			bsShowMessage("Digite uma Idade Final ou apage a Idade Inicial","I")
      		Exit Sub
		End If
	Else
		If Not CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
			bsShowMessage("Digite uma Idade Inicial ou apage a Idade Final","I")
      		Exit Sub
		End If
	End If

	CIDsArray() = Split(CurrentQuery.FieldByName("CID").AsString, "|_|", -1)


	If Not CurrentQuery.FieldByName("IDADEINICIAL").IsNull Then
		vsCampos = vsCampos + "IDADEINICIAL,"
		vsParametros = vsParametros + ":IDADEINICIAL,"
	End If
	If Not CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
		vsCampos = vsCampos + "IDADEFINAL,"
		vsParametros = vsParametros + ":IDADEFINAL,"
	End If


	Set qInsert = NewQuery
	qInsert.Active = False
	qInsert.Clear
	qInsert.Add("INSERT                                                                      ")
	qInsert.Add("  INTO ANS_SIP_ANEXO_ITEM_CID                                               ")
	qInsert.Add("       (HANDLE, SIPANEXO," + vsCampos + "TIPOPERIODO, SEXO, CID)            ")
	qInsert.Add("VALUES (:HANDLE, :SIPANEXO," + vsParametros + " :TIPOPERIODO, :SEXO, :CID)  ")

	For i = 0 To UBound(CIDsArray)
   		If CIDsArray(i) <> "" Then
			qInsert.ParamByName("HANDLE").AsInteger = NewHandle("ANS_SIP_ANEXO_ITEM_CID")
			qInsert.ParamByName("SIPANEXO").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO_ITEM")

			If Not CurrentQuery.FieldByName("IDADEINICIAL").IsNull Then
				qInsert.ParamByName("IDADEINICIAL").AsInteger = CurrentQuery.FieldByName("IDADEINICIAL").AsInteger
			End If

			If Not CurrentQuery.FieldByName("IDADEFINAL").IsNull Then
				qInsert.ParamByName("IDADEFINAL").AsInteger = CurrentQuery.FieldByName("IDADEFINAL").AsInteger
			End If

			qInsert.ParamByName("TIPOPERIODO").AsString = CurrentQuery.FieldByName("TIPOPERIODO").AsString
			qInsert.ParamByName("SEXO").AsString = CurrentQuery.FieldByName("Sexo").AsString
			qInsert.ParamByName("CID").AsInteger = CLng(CIDsArray(i))
			qInsert.ExecSQL
    	End If
	Next
	Set qInsert = Nothing

End Sub
