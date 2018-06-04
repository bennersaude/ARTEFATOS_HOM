'HASH: 80F1DBDFC0DDC409B1A0964922FF0456
 
 
Public Sub TABLE_BeforePost(CanContinue As Boolean) 
  If (CurrentQuery.FieldByName("TERMINO").AsDateTime <= CurrentQuery.FieldByName("INICIO").AsDateTime) Then 
    If (WebMode) Then 
      CancelDescription = "Horário final tem que ser maior que horário inicial" 
    Else 
      MsgBox("Horário final tem que ser maior que horário inicial") 
    End If 
    CanContinue = False 
  End If 
End Sub 
