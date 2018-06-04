'HASH: D577E66001D14DE2B2EB7785D358FD54
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >= CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
      bsShowMessage("A data inicial deve ser inferior à data final!", "E")
      CanContinue = False
  End If
End Sub
