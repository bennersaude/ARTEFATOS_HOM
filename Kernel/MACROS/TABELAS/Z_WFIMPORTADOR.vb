'HASH: 0E3060BFF9F123CD4F5454EBBBB5CBE3
Public Sub CAMINHO_OnBtnClick() 
	CurrentQuery.FieldByName("CAMINHO").AsString = OpenDialog 
End Sub 
 
Public Sub CAMINHOBUILDER_OnBtnClick() 
	CurrentQuery.FieldByName("CAMINHOBUILDER").AsString = OpenDialog 
End Sub 
