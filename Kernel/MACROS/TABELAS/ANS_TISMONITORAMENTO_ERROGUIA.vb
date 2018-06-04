'HASH: 937921A4A06DF6D51566782AB538BDCC

Public Sub BOTAOCORRIGIRCAMPO_OnClick()
  Dim form As CSVirtualForm
  Set form = NewVirtualForm

  SessionVar("CODIGOCAMPOCOMERRO") = CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString

  If CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString = "062" Or CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString = "063" Then

    Dim qSql As BPesquisa
    Set qSql = NewQuery
    qSql.Add("SELECT CODIGOGLOSATISS      ")
    qSql.Add("  FROM SAM_MOTIVOGLOSA_TISS ")
    qSql.Add(" WHERE HANDLE = :HANDLE     ")

    qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CODIGOERRO").AsInteger
    qSql.Active = True

    If qSql.FieldByName("CODIGOGLOSATISS").AsInteger = 5034 Then
      form.Caption = "Inclusão de declaração"
      form.TableName = "ANS_TISMONITORAMENTO_GUIA_DEC"
      form.Physical = True
      form.Height = 200
      form.Width = 300
      form.Show
    End If

    qSql.Active = False
    Set qSql = Nothing
  Else
    form.Caption = "Correção da Guia"
    form.TableName = "ANS_TISMONITORAMENTO_GUIA"
    form.CurrentHandle = CurrentQuery.FieldByName("GUIA").AsInteger
    form.Physical = True
    form.Height = 600
    form.Width = 600
    form.Show
  End If

  Set form = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  SessionVar("CODIGOCAMPOCOMERRO") = CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString
End Sub
