'HASH: 4B32BED21DB42370023ACBF7247201AF

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Obj As Object
  Set Obj = CreateBennerObject("Calc.BCalc")
  CanContinue = Obj.DeleteSheet(CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing
End Sub

