'HASH: 597C30B67802BAC8E4F8020A091627DD
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
     BOTAOGERAREVENTOS.Visible=False 
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("CODIGO").AsInteger = 0 Then
		bsShowMessage("O código deve ser superior a zero", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
