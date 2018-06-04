'HASH: 1A414B044C266C93232D0BA34F9CD4F2

Public Sub EXPORTAR_OnClick()
  Dim obj As Object
  Set obj = CreateBennerObject("CS.ImageImpExp")
  obj.Prepare(CurrentSystem)
  obj.ExportFile(CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set obj = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "EXPORTAR" Then
		EXPORTAR_OnClick
	End If
End Sub
