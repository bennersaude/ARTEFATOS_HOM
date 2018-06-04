'HASH: CB4736FF1A97589AC4437F0E9F803FE6
'MACRO: CA_ATEND_DOCANEXO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.InInsertion Or CurrentQuery.InEdition Then
     If (CurrentQuery.FieldByName("DATAENTREGA").AsDateTime > ServerNow) Then
    	bsShowMessage("A data de entregra não deve ser maior que a data atual!", "E")
    	CanContinue = False
    	Exit Sub
  	 End If
  End If
End Sub
