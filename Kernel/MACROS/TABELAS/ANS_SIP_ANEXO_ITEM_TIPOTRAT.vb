'HASH: 73A1F38AB852B8E627371A1FC932D4EA
'Macro: ANS_SIP_ANEXO_ITEM_TIPOTRAT
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM ANS_SIP_ANEXO_ITEM_TIPOTRAT WHERE SIPANEXO = :SIPANEXO AND TIPOTRATAMENTO = :TIPOTRATAMENTO AND HANDLE <> :HANDLE")
  sql.ParamByName("SIPANEXO").AsInteger = CurrentQuery.FieldByName("SIPANEXO").AsInteger
  sql.ParamByName("TIPOTRATAMENTO").AsInteger = CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If Not sql.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Tipo de Tratamento já cadastrado para esse ítem.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
