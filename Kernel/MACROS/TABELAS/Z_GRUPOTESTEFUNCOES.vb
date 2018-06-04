'HASH: C510DCBDEA9A3DB3A9F35DBBB23C8435
Option Explicit 
 
Sub ShowXML(campo As String) 
Dim obj As Object 
  Set obj = CreateBennerObject("CS.CustomFunctions") 
  obj.ShowContainerXML(CurrentQuery.FieldByName(campo).AsString) 
  Set obj = Nothing 
End Sub 
 
Public Sub BOTAOSIMULAR_OnClick() 
 
Dim obj As Object 
  Set obj = CreateBennerObject("CS.CustomFunctions") 
  obj.ShowSimulation(CurrentQuery.FieldByName("NOME").AsString, CurrentQuery.FieldByName("ENTRADA").AsString, CurrentQuery.FieldByName("SAIDA").AsString) 
  Set obj = Nothing 
End Sub 
 
Public Sub BOTAOVERENTRADA_OnClick() 
  ShowXML "ENTRADA" 
End Sub 
 
Public Sub BOTAOVERSAIDA_OnClick() 
  ShowXML "SAIDA" 
End Sub 
