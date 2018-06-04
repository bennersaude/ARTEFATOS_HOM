'HASH: 0E4C5D2564D90757CCE76703725139C9
 
 
Public Sub BOTAOGRAVAR_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("cscommon.rotinas") 
  obj.BeginRecordFunctions(CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set obj = Nothing 
End Sub 
 
Public Sub BOTAOPARAR_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("cscommon.rotinas") 
  obj.StopRecordFunctions 
  Set obj = Nothing 
End Sub 
 
Public Sub BOTAOTESTAR_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("cscommon.rotinas") 
  obj.TestFunctions(CurrentQuery.FieldByName("HANDLE").AsInteger) 
  Set obj = Nothing 
  MsgBox "Teste finalizado." 
End Sub 
