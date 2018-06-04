'HASH: 767DA44D4404D02F2C995E8C23DA6A21
Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qInsert As Object
	Dim i As Long

	Set qInsert = NewQuery

	Dim ItensArray() As String

	ItensArray() = Split(CurrentQuery.FieldByName("EVENTOS").AsString, "|_|", -1)

	qInsert.Active = False
	qInsert.Clear
	qInsert.Add("INSERT                                  ")
	qInsert.Add("  INTO ANS_SIP_ANEXO_ITEM_EVENTO        ")
	qInsert.Add("       (HANDLE, SIPITEM, EVENTO)        ")
	qInsert.Add("VALUES (:HANDLE, :ITEM, :ITEMAPROPRIADO)")

	For i = 0 To UBound(ItensArray)
   		If ItensArray(i) <> "" Then
			qInsert.ParamByName("HANDLE").AsInteger = NewHandle("ANS_SIP_ANEXO_ITEM_EVENTO")
			qInsert.ParamByName("ITEM").AsInteger = RecordHandleOfTable("ANS_SIP_ANEXO_ITEM")
			qInsert.ParamByName("ITEMAPROPRIADO").AsInteger = CLng(ItensArray(i))
			qInsert.ExecSQL
    	End If
	Next
	Set qInsert = Nothing
End Sub
