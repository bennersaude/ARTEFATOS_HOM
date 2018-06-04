'HASH: D5668F8A24FAC4835D225D2DA11F0459

Public Sub BOTAOALTERARDATAGUIA_OnClick()

  Dim qSql As BPesquisa
  Set qSql = NewQuery
  qSql.Add("SELECT ROTINAMONITORAMENTOGUIA        ")
  qSql.Add("  FROM ANS_TISMONITORAMENTO_GUIA_PROC ")
  qSql.Add(" WHERE HANDLE = :HANDLE               ")

  qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PROCEDIMENTO").AsInteger
  qSql.Active = True

  SessionVar("CODIGOCAMPOCOMERRO") = CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString

  Dim form As CSVirtualForm
  Set form = NewVirtualForm

  form.Caption = "Correção da Guia"
  form.TableName = "ANS_TISMONITORAMENTO_GUIA"
  form.CurrentHandle = qSql.FieldByName("ROTINAMONITORAMENTOGUIA").AsInteger
  form.Physical = True
  form.Height = 600
  form.Width = 600
  form.Show

  qSql.Active = False

  Set qSql = Nothing
  Set form = Nothing
End Sub

Public Sub BOTAOCORRIGIRCAMPO_OnClick()
  Dim form As CSVirtualForm
  Set form = NewVirtualForm

  Dim qBuscaGlosa As BPesquisa
  Set qBuscaGlosa = NewQuery
  qBuscaGlosa.Add("SELECT CODIGOGLOSATISS      ")
  qBuscaGlosa.Add("  FROM SAM_MOTIVOGLOSA_TISS ")
  qBuscaGlosa.Add(" WHERE HANDLE = :PHANDLE    ")

  qBuscaGlosa.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("CODIGOERRO").AsInteger
  qBuscaGlosa.Active = True

  SessionVar("CODIGOCAMPOCOMERRO") = CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString
  SessionVar("CODIGOGLOSA") = qBuscaGlosa.FieldByName("CODIGOGLOSATISS").AsString

  qBuscaGlosa.Active = False
  Set qBuscaGlosa = Nothing

  form.Caption = "Correção do Procedimento"
  form.TableName = "ANS_TISMONITORAMENTO_GUIA_PROC"
  form.CurrentHandle = CurrentQuery.FieldByName("PROCEDIMENTO").AsInteger
  form.Physical = True
  form.Height = 600
  form.Width = 600
  form.Show

  Set form = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  SessionVar("CODIGOCAMPOCOMERRO") = CurrentQuery.FieldByName("IDENTIFICADORCAMPO").AsString
End Sub
