'HASH: 2554FB096BECFDAB80E3A807C869A595
'#Uses "*bsShowMessage"
'CLI_ROTGUIAFALTAAGENDA

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim rot As Object
  Set rot = NewQuery
  rot.Clear
  rot.Add("SELECT DATAPROCESSAMENTO FROM CLI_ROTGUIAFALTA WHERE HANDLE = :ROTGUIAFALTA")
  rot.ParamByName("ROTGUIAFALTA").AsInteger = CurrentQuery.FieldByName("ROTGUIAFALTA").AsInteger
  rot.Active = True
  If Not rot.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("A rotina já foi processada!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set rot = Nothing
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("A situação está pendente!", "E")
    CanContinue = False
    Exit Sub
  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString = "J" And CurrentQuery.FieldByName("JUSTIFICATIVA").IsNull Then
    bsShowMessage("É necessário uma justificativa!", "E")
    CanContinue = False
    Exit Sub
  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString = "J" Then
    CurrentQuery.FieldByName("IGNORAPROCESSAMENTO").AsString = "S"
  End If
  CurrentQuery.FieldByName("DATA").AsDateTime = ServerNow
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
End Sub

