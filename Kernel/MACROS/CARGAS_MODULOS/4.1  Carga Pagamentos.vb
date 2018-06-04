'HASH: 2941E04D4EA75AAEDC6D125E247FD787


Public Sub COPIARPRECOS_OnClick()
  Dim Interface As Object

  Set Interface =CreateBennerObject("BSPRE001.Rotinas")
  Interface.CopiarPrecos(CurrentSystem)
  Set Interface =Nothing
End Sub
