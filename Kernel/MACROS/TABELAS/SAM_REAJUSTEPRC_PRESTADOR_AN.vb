'HASH: DC75CD289711660E868450CE4C274FFF
'Macro: SAM_REAJUSTEPRC_PRESTADOR_AN
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
  Dim TIPO As Object
  Set TIPO = NewQuery
  TIPO.Add("SELECT HANDLE FROM SAM_REAJUSTEPRC_PARAMTIPO T WHERE T.REAJUSTEPRCPARAM = :PARAM AND ")
  TIPO.Add("T.TIPODOREAJUSTE = 'A'")
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

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem,"E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem,"A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

