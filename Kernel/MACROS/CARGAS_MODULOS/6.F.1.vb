'HASH: 998343362B8C208529CBF01D356FC0B1
Public Sub AGENDARPEGS_OnClick()
  Dim dll As Object
  Set dll =CreateBennerObject("sampeg.processar")
  dll.SelecionarPegsAgenda(CurrentSystem)
  Set dll =Nothing
End Sub
 
