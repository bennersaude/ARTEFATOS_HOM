'HASH: EF8F302C4BA52D569F819587B9C0116D
'Macro: ANS_SIP_ANEXO_ITEM_REGIMEATEND
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM ANS_SIP_ANEXO_ITEM_REGIMEATEND WHERE SIPITEM = :SIPITEM AND REGIMEATENDIMENTO = :REGIMEATENDIMENTO AND HANDLE <> :HANDLE")
  sql.ParamByName("SIPITEM").AsInteger = CurrentQuery.FieldByName("SIPITEM").AsInteger
  sql.ParamByName("REGIMEATENDIMENTO").AsInteger = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If Not sql.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Regime de Atendimento já cadastrado para esse ítem.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
