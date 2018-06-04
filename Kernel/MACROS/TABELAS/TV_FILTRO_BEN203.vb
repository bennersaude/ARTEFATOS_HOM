'HASH: 86135F6DCF265E43B755DDB297D2E72D
 
'#Uses "*bsShowMessage


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("TABFILTRO").AsInteger = 1 Or CurrentQuery.FieldByName("TABFILTRO").AsInteger = 3 Then
	If CurrentQuery.FieldByName("DATAINTINI").AsDateTime > CurrentQuery.FieldByName("DATAINTFIN").AsDateTime Then
       bsShowMessage("Data Inicial superior a data final","E")
       CanContinue = False
	End If
  End If

End Sub
