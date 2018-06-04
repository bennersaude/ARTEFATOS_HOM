'HASH: 2841F9260D7EA22CA74A51918F19A9D3
Public Sub TABLE_AfterEdit()
	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_AfterInsert()
	CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub
