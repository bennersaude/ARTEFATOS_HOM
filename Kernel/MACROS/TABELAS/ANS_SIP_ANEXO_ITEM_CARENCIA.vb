'HASH: 1D8AB8078D6B64215647AE25DCCFBC4F
'Macro: ANS_SIP_ANEXO_ITEM_CARENCIA
'#Uses "*bsShowMessage"

 Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM ANS_SIP_ANEXO_ITEM_CARENCIA WHERE SIPANEXO = :SIPANEXO AND CARENCIA = :CARENCIA AND HANDLE <> :HANDLE")
  sql.ParamByName("SIPANEXO").AsInteger = CurrentQuery.FieldByName("SIPANEXO").AsInteger
  sql.ParamByName("CARENCIA").AsInteger = CurrentQuery.FieldByName("CARENCIA").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If Not sql.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Carência já cadastrada para esse ítem.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
