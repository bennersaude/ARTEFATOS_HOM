'HASH: B5182ACDC3949B649FD37E3CA022C947
 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("PRESTADOR").AsInteger > 0 And CurrentQuery.FieldByName("PESSOA").AsInteger > 0) _
	  Or (CurrentQuery.FieldByName("PRESTADOR").AsInteger > 0 And CurrentQuery.FieldByName("BENEFICIARIO").AsInteger > 0) _
	  Or (CurrentQuery.FieldByName("PESSOA").AsInteger > 0 And CurrentQuery.FieldByName("BENEFICIARIO").AsInteger > 0) Then
	  	bsShowMessage("Escolha somente um tipo de responsável", "E")
	  	CanContinue = False
	  	Exit Sub
	End If
End Sub
