'HASH: 4E19208F618671D27ACFE68BB0EC79F4

Public Sub IMPORTAR_OnClick()
Dim obj As Object
  Set obj =CreateBennerObject("CS.ImageImpExp")
  obj.Prepare(CurrentSystem)
  obj.ImportFile
  Set obj =Nothing
End Sub
