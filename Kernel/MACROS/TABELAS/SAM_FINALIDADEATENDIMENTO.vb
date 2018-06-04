'HASH: 9CE7C3AE4F9094EB3C10D1F0ADD67418
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qSel As Object
  Set qSel = NewQuery
  qSel.Active = False
  qSel.Clear
  qSel.Add("SELECT HANDLE FROM SAM_FINALIDADEATENDIMENTO WHERE CODIGO = :CODIGO AND HANDLE <> :HANDLE")
  qSel.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  qSel.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSel.Active = True

  If Not qSel.FieldByName("HANDLE").IsNull Then
    MsgBox("Código já cadastrado.")
    CanContinue = False
    Set qSel = Nothing
    Exit Sub
  End If

  Set qSel = Nothing

End Sub
