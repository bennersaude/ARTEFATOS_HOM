'HASH: C25A07170EA2643A49F3B16EE5C95C28
 Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qInsert As Object
	Dim i As Long

	Set qInsert = NewQuery

	Dim ItensArray() As String

	ItensArray() = Split(CurrentQuery.FieldByName("GRAU").AsString, "|_|", -1)

	qInsert.Active = False
	qInsert.Clear
	qInsert.Add("INSERT                                    ")
	qInsert.Add("  INTO ANS_SIP_ANEXO_ITEM_GRAU            ")
	qInsert.Add("       (HANDLE, SIPITEM, GRAU)            ")
	qInsert.Add("VALUES (:HANDLE, :ITEM, :ITEMAPROPRIADO)  ")

	For i = 0 To UBound(ItensArray)
   		If ItensArray(i) <> "" Then
			qInsert.ParamByName("HANDLE").AsInteger = NewHandle("ANS_SIP_ANEXO_ITEM_GRAU")
			qInsert.ParamByName("ITEM").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO_ITEM")
			qInsert.ParamByName("ITEMAPROPRIADO").AsInteger = CLng(ItensArray(i))
			qInsert.ExecSQL

    	End If
	Next
	Set qInsert = Nothing

End Sub
