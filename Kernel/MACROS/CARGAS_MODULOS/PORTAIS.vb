'HASH: C183705587BBE5A00D846EF3EB61AD40
 
Public Sub VISOES_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebVisionDesigner") 
  Obj.Exec(CurrentSystem, 0) 
  Set Obj = Nothing 
End Sub 
