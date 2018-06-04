'HASH: D1D73ED902B093EFEFA2281157FABCF1
'Macro: SAM_AGENTEVENDAS
'#uses "*bsShowMessage"
Public Sub TABLE_AfterScroll()
  CurrentQuery.FieldByName("CPF").Mask= "999\.999\.999\-99;0;_"
End Sub
Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  CurrentQuery.FieldByName("CPF").Mask= ""
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("CPF").AsString <> "" Then
	  If Not IsValidCPF(CurrentQuery.FieldByName("CPF").AsString) Then
	    bsShowMessage("CPF Inválido", "I")
	    CPF.SetFocus
	  End If
  End If
  TABLE_AfterScroll
End Sub

