'HASH: FFD9B97F326CB10F30FAF722C8CDEBAA

'CLI_ATIVIDADE

'#Uses "*ProcuraGrauValido"
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("EVENTO").IsNull Then
    If CurrentQuery.FieldByName("GRAU").IsNull Then
      bsShowMessage("Se for escolhido um evento, o grau passa a ser obrigatório!", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  If CurrentQuery.FieldByName("PRONTOATENDIMENTO").AsString = "S" Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT COUNT(*) TOTAL FROM CLI_ATIVIDADE WHERE PRONTOATENDIMENTO = 'S' AND HANDLE <> :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If SQL.FieldByName("TOTAL").AsInteger > 0 Then
      bsShowMessage("Só pode existir uma atividade de urgência!", "E")
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("EVENTO").IsNull Then
      bsShowMessage("Na atividade de urgência é obrigatório indicar o evento!", "E")
      CanContinue = False
      Exit Sub
    End If
    If CurrentQuery.FieldByName("GRAU").IsNull Then
      bsShowMessage("Na atividade de urgência é obrigatório indicar o grau!", "E")
      CanContinue = False
      Exit Sub
    End If
    Set SQL = Nothing
  End If
End Sub


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  ShowPopup = False
  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    Exit Sub
  End If

  Dim vHandle As Long
  vHandle = ProcuraGrauValido(CurrentQuery.FieldByName("EVENTO").AsInteger, GRAU.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
End Sub

