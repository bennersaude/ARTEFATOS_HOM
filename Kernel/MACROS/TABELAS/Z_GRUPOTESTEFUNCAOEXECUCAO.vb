'HASH: 039045262466C47876B1FA42F0919ABF
 
 
Public Sub BOTAOVERSAIDA_OnClick() 
Dim obj As Object 
  Set obj = CreateBennerObject("CS.CustomFunctions") 
  obj.ShowContainerXML(CurrentQuery.FieldByName("SAIDA").AsString) 
  Set obj = Nothing 
End Sub 
