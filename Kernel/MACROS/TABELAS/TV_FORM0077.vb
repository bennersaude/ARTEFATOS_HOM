'HASH: 4156EFC8DA6BD5E45CA22A7B645E9F89
 '#Uses "*bsShowMessage"
 Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qInsert As Object
	Dim i As Long

	Set qInsert = NewQuery

	Dim ItensArray() As String

	ItensArray() = Split(CurrentQuery.FieldByName("CLASSES").AsString, "|_|", -1)

	qInsert.Active = False
	qInsert.Clear
	qInsert.Add("INSERT                                    ")
	qInsert.Add("  INTO ANS_SIP_ANEXO_ITEM_CLASSE          ")
	qInsert.Add("       (HANDLE, SIPITEM, CLASSEGERENCIAL) ")
	qInsert.Add("VALUES (:HANDLE, :ITEM, :ITEMAPROPRIADO)  ")

	For i = 0 To UBound(ItensArray)
   		If ItensArray(i) <> "" Then
			qInsert.ParamByName("HANDLE").AsInteger = NewHandle("ANS_SIP_ANEXO_ITEM_CLASSE")
			qInsert.ParamByName("ITEM").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO_ITEM")
			qInsert.ParamByName("ITEMAPROPRIADO").AsInteger = CLng(ItensArray(i))
			qInsert.ExecSQL

    	End If
	Next

	Set qInsert = Nothing
End Sub
