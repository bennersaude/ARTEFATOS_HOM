'HASH: F489A460BFC28FB10C9DAFE5C2D3C504
Public Sub MODULE_OnEnter() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.Updater") 
  Obj.Exec(CurrentSystem) 
  Set Obj = Nothing 
End Sub 
 
Public Sub SERVICOS_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Benner.Tecnologia.Architect.CustomServicesForm") 
  Obj.Exec(CurrentSystem) 
  Set Obj = Nothing 
End Sub 
 
Public Sub VISOES_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebVisionDesigner") 
  Obj.Exec(CurrentSystem, 0) 
  Set Obj = Nothing 
End Sub 
