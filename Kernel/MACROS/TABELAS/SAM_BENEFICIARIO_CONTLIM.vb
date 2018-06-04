'HASH: 433BB0C4AE1E2402FB72B315048C138F
Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("CONTLIMMIGRADA").AsInteger > 0 Then
    CONTLIMMIGRADA.Visible = True
    TABTIPOLIMITE.Visible = False
  Else
    CONTLIMMIGRADA.Visible = False
    TABTIPOLIMITE.Visible = True
  End If
End Sub
