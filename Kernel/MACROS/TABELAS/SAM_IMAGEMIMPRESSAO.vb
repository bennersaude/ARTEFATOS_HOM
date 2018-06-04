'HASH: 8CD82E97A39D344C9976475916A94638

Public Sub BOTAOIMPORTAR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamImpressao.Rotinas")
  Obj.Importar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
End Sub

