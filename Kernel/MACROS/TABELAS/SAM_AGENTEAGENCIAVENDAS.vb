'HASH: F829D524D4D9DE03ED38EFD8F2E80E44
'Macro: SAM_AGENTEAGENCIAVENDAS
'#Uses "*bsShowMessage"
Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If (VisibleMode And NodeInternalCode <> 987123) Or (WebMode And WebVisionCode <> "V_SAM_AGENTEAGENCIAVENDAS") Then 'A pasta com o código 987123 é somente leitura e, portanto, a data final não pode ficar em edição
    If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
      DATAFINAL.ReadOnly=False
	Else
	  DATAFINAL.ReadOnly=True	  
	End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And _
     (CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) Then
     CanContinue = False
     bsShowMessage("A Data final, se informada, deve ser maior ou igual a inicial", "E")
     Exit Sub
  Else
     CanContinue = True
  End If  

End Sub
