'HASH: E943FD7A5CD480021F05B69389D541A5
Dim obj

Private Sub OpenForm()
  Set obj = CreateBennerObject("CS.FORMULARIOS")
  obj.CreateObjForm("Formulário")
End Sub


Private Sub CloseForm()
  obj.FreeObjForm
  Set obj = Nothing
End Sub


Public Sub EXCLUIR_OnClick()
  OpenForm
  obj.DeleteForm(CurrentQuery.FieldByName("HANDLE").AsInteger)
  CloseForm
  RefreshNodesWithTable("Z_FORMULARIOS")
End Sub


Public Sub EXPORTAR_OnClick()
  OpenForm
  obj.SaveToFile(CurrentQuery.FieldByName("HANDLE").AsInteger)
  CloseForm
End Sub


Public Sub IMPORTAR_OnClick()
  OpenForm
  obj.LoadFromFile
  CloseForm
  RefreshNodesWithTable("Z_FORMULARIOS")
End Sub

Public Sub TESTFORM_OnClick()
  OpenForm
  obj.TestForm(CurrentQuery.FieldByName("HANDLE").AsInteger)
  CloseForm
End Sub

