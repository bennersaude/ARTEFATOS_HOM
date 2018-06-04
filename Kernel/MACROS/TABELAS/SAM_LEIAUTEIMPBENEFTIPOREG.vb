'HASH: F332E54844E9FAC6BEF166244976E7C4
'#Uses "*bsShowMessage"

Dim vgtipo  As String

Public Sub TABLE_AfterEdit()
  vgtipo = CurrentQuery.FieldByName("TIPOREGISTRO").AsString

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.State <> 3 And vgtipo <> CurrentQuery.FieldByName("TIPOREGISTRO").AsString Then
    Dim Sql As Object
    Set Sql = NewQuery
    Sql.Add("SELECT COUNT(1) NREC FROM SAM_LEIAUTEIMPBENEFTIPOREGCAMP WHERE TIPOREGISTRO = :HANDLE ")
    Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    Sql.Active=True
    If Sql.FieldByName("NREC").AsInteger > 0 Then
      bsShowMessage("Não permitido alterar tipo registro quando existir campo cadastrado!", "E")
      CanContinue=False
      Exit Sub
      Set Sql = Nothing
    End If
    Set Sql = Nothing
  End If


  If CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "A" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "00"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "B" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "01"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "C" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "15"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "D" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "16"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "E" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "17"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "F" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "18"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "G" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "20"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "H" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "30"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "I" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "31"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "J" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "32"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "K" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "40"
  ElseIf CurrentQuery.FieldByName("TIPOREGISTRO").AsString = "Z" Then
    CurrentQuery.FieldByName("CODIGO").AsString = "99"
  End If


End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("CODIGO").AsString = "00"
End Sub
