'HASH: 2D30AEFC51301EF3C09223DF7D3BA5AB
Public Sub EDITAR_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebVisionDesigner") 
  Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub PUBLICAR_OnClick() 
Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebPublisher") 
  Obj.PublishVision(CurrentSystem, -1, CurrentQuery.FieldByName("NOME").AsString) 
  Set Obj = Nothing 
End Sub 
 
Public Sub SEGURANCA_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.WebSecurity") 
  Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
 
Public Sub SUBSTITUTAS_OnClick() 
  Dim Obj As Object 
  Set Obj = CreateBennerObject("Pyxis.PermissionConfig") 
  Obj.Vision(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set Obj = Nothing 
End Sub 
