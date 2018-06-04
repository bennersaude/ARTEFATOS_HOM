'HASH: E39E923204AAA4DE5E7A2EADD6402819
'Macro: ANS_SIP_ANEXO_ITEM_ESPECIALIDA
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM ANS_SIP_ANEXO_ITEM_ESPECIALIDA WHERE SIPANEXO = :SIPANEXO AND ESPECIALIDADE = :ESPECIALIDADE AND HANDLE <> :HANDLE")
  sql.ParamByName("SIPANEXO").AsInteger = CurrentQuery.FieldByName("SIPANEXO").AsInteger
  sql.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If Not sql.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Especialidade já cadastrada para esse ítem.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
