'HASH: 994E6903677B506127873745021D0738
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qBusca As Object
  Set qBusca = NewQuery

  'Verifica se a ordem digitada já está sendo utilizada
  qBusca.Clear
  qBusca.Add("SELECT COUNT(TIPOFATURA) QTDE FROM SFN_TIPODOCUMENTO_BAIXA ")
  qBusca.Add("WHERE ORDEM =:ORDEM                                        ")
  qBusca.Add("  AND HANDLE <>:HANDLE                                     ")
  qBusca.Add("  AND TIPODOCUMENTO =:HTIPODOC                             ")
  qBusca.ParamByName("ORDEM").AsString = CurrentQuery.FieldByName("ORDEM").Value
  qBusca.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qBusca.ParamByName("HTIPODOC").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").Value
  qBusca.Active = True

  If qBusca.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Ordem já cadastrada!", "I")
    ORDEM.SetFocus
    qBusca.Active = False
    Set qBusca = Nothing
    CanContinue = False
    Exit Sub
  End If

  'Verifica se já existe um tipo de fatura cadastrada
  qBusca.Clear
  qBusca.Add("SELECT COUNT(TIPOFATURA) QTDE FROM SFN_TIPODOCUMENTO_BAIXA ")
  qBusca.Add("WHERE TIPODOCUMENTO =:HTIPODOC                             ")
  qBusca.Add("  AND TIPOFATURA =:HTIPOFAT                                ")
  qBusca.Add("  AND HANDLE <>:HANDLE                                     ")
  qBusca.ParamByName("HTIPODOC").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").Value
  qBusca.ParamByName("HTIPOFAT").AsInteger = CurrentQuery.FieldByName("TIPOFATURA").Value
  qBusca.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").Value
  qBusca.Active = True

  If qBusca.FieldByName("QTDE").AsInteger > 0 Then
    bsShowMessage("Tipo de fatura já cadastrada!", "I")
    TIPOFATURA.SetFocus
    qBusca.Active = False
    Set qBusca = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set qBusca = Nothing

End Sub

