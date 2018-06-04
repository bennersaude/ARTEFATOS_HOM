'HASH: 20A366D8BD412908AB20BBD5819C42C6
'Macro: SAM_REAJUSTEPRC_PRESTADOR_SL
'#Uses "*bsShowMessage"

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

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

