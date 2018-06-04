'HASH: 7F70E1E4E3BC7F81C2B82BC5DE480A27
'Macro: SFN_COMPETFIN
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  RefreshNodesWithTable("SFN_COMPETFIN")
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qBusca As Object
  Set qBusca = NewQuery

    qBusca.Clear
    qBusca.Add("SELECT HANDLE                 ")
    qBusca.Add("  FROM SFN_COMPETFIN          ")
    qBusca.Add(" WHERE TIPOFATURAMENTO =:TFAT ")
    qBusca.Add("   AND COMPETENCIA =:COMPET   ")
    qBusca.Add("   AND HANDLE <>:HANDLE       ")
    qBusca.ParamByName("TFAT").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
    qBusca.ParamByName("COMPET").AsDateTime = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
    qBusca.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qBusca.Active = True

    If qBusca.FieldByName("HANDLE").AsInteger > 0 Then
      bsShowMessage("Competência já cadastrada!", "E")
      COMPETENCIA.SetFocus
      CanContinue = False
    End If

  Set qBusca = Nothing
End Sub
