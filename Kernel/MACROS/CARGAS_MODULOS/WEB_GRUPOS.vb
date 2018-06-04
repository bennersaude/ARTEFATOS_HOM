'HASH: 0B67FD411CA96AB57F4CCC89FB4D0420
Public Sub PUBLICAR_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishGroups(CurrentSystem, -1) 
  Set Obj = Nothing 
End Sub 
