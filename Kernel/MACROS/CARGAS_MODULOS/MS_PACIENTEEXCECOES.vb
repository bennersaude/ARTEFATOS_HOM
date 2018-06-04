'HASH: E9F8FC4B065025BDFD84D45456F68D9D
 


Public Sub ATRIBUIREXAME_OnClick()
  Dim obj As Object
  Set obj =CreateBennerObject("BSMED001.AtribuirPaciente")
  obj.Exec(CurrentSystem)
  Set obj =Nothing
End Sub
