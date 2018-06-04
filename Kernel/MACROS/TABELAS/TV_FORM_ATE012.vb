'HASH: FD83D789D7CA2E0B24DEE6A4157D2F07
Option Explicit

Public Sub TABLE_AfterInsert()

	If UserVar("FILTRO_TV_FORM_ATE012") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FORM_ATE012"),CurrentQuery.TQuery)
	End If

End Sub

Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FORM_ATE012") = DatasetToXML(CurrentQuery.TQuery,"")

End Sub
