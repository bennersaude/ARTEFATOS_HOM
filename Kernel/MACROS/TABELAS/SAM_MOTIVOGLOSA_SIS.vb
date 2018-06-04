'HASH: 2D95B18DB7BB2EA03CCC8F70BA09B3B5

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qExiste As Object

  Set qExiste = NewQuery

  qExiste.Clear
  qExiste.Add("SELECT COUNT(1) QTDE")
  qExiste.Add("  FROM SAM_MOTIVOGLOSA_SIS")
  qExiste.Add(" WHERE SISMOTIVOGLOSA = :SISMOTIVOGLOSA")
  qExiste.Add("   AND HANDLE <> :HANDLE")
  qExiste.ParamByName("SISMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("SISMOTIVOGLOSA").AsInteger
  qExiste.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qExiste.Active = True

  If qExiste.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Glosa de sistema já cadastrada, não permitido!", "I")
    CanContinue = False
  End If

  Set qExiste = Nothing

End Sub
