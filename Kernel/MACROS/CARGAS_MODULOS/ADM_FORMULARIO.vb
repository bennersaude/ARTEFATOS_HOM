'HASH: 4AC47B04A200C2B0804C6C3A500A0BC4
Public Sub IMPORTAR_OnClick()
Dim obj
  Set obj =CreateBennerObject("CS.FORMULARIOS")
  obj.CreateObjForm("Formulários")
  obj.LoadFromFile
  obj.FreeObjForm
  Set obj =Nothing
  RefreshNodesWithTable("Z_FORMULARIOS")
End Sub
