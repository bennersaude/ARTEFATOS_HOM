'HASH: 29E2DEB6F6BC4509D9126F85B3A72736
'Macro: SAM_REAJUSTEPRC_REDE_SL

Public Sub TABLE_AfterInsert()
  Dim TIPO As Object
  Set TIPO = NewQuery
  TIPO.Add("SELECT HANDLE FROM SAM_REAJUSTEPRC_PARAMTIPO T WHERE T.REAJUSTEPRCPARAM = :PARAM AND ")
  TIPO.Add("T.TIPODOREAJUSTE = 'S'")
  TIPO.ParamByName("PARAM").Value = CurrentQuery.FieldByName("REAJUSTEPRCPARAM").AsInteger
  TIPO.Active = True
  If TIPO.EOF Then
    SetParamTipo = False
  Else
    setParamTipo = True
    CurrentQuery.FieldByName("PARAMTIPO").Value = TIPO.FieldByName("HANDLE").AsInteger
  End If
  TIPO.Active = False
End Sub


