'HASH: 9B32C3E6B7870A3111204B2693B04DF4
Public Sub EDITAR_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebVisionDesigner") 
  Obj.Exec(CurrentSystem, 0) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICAR_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishVisions(CurrentSystem, -1) 
  Set Obj = Nothing 
End Sub 
