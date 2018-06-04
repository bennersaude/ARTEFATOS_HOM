'HASH: 24ACE31CB410CE24E1B0316ECA074062
'Macro: ANS_SIP_ANEXO_ITEM_CLASSE
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM ANS_SIP_ANEXO_ITEM_CLASSE WHERE SIPITEM = :SIPITEM AND CLASSEGERENCIAL = :HCLASSE AND HANDLE <> :HANDLE")
  sql.ParamByName("SIPITEM").AsInteger = CurrentQuery.FieldByName("SIPITEM").AsInteger
  sql.ParamByName("HCLASSE").AsInteger = CurrentQuery.FieldByName("CLASSEGERENCIAL").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If Not sql.FieldByName("HANDLE").IsNull Then
  	bsShowMessage("Classe Gerencial já cadastrada para esse ítem.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
