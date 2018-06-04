'HASH: 21DB8791771ABBF11CF0B018D0E63A8B
 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  If (CurrentQuery.FieldByName("NAOANEXAR").AsString = "S") And (CurrentQuery.FieldByName("NAOSALVAR").AsString = "S") Then 
    CanContinue = False 
    MsgBox("O relatório deve ser salvo ou enviado por e-mail!", vbInformation) 
  End If 
End Sub 
