'HASH: C2CED48ADFAD741982D98C822E4414A4
'#uses "*bsShowMessage"
Dim InsertQuery As Object

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim vHandleAux As Long
  Dim vReembolso As Boolean
  Dim vNull As Integer

  Set SQL = NewQuery
  Set InsertQuery = NewQuery

  InsertQuery.Clear
  InsertQuery.Add("INSERT INTO SAM_PREACERTO_GUIA(HANDLE, GUIA, PEG, DATAATENDIMENTO, USUARIOGERACAO, DATAGERACAO, ACERTOREALIZADO,")
  InsertQuery.Add("                               DATACREDITOBENEFICIARIO, DATADEBITOBENEFICIARIO,")
  InsertQuery.Add("                               DATACREDITORECEBEDOR, DATADEBITORECEBEDOR,")
  InsertQuery.Add("                               ACERTOLOTE, TABCOBPAGACERTOSBENEF,")
  InsertQuery.Add("                               AUTORIZACAO, EXECUTOR, BENEFICIARIO, LOCALEXECUCAO, BANCO, AGENCIA, CONTACORRENTENOME,")
  InsertQuery.Add("                               CONTACORRENTENUMERO, CONTACORRENTEDV, CONTACORRENTECPFCNPJ, RECEBEDOR)")
  InsertQuery.Add("                    VALUES (:HANDLE, :GUIA, :PEG, :DATAATENDIMENTO, :USUARIOGERACAO, :DATAGERACAO, :ACERTOREALIZADO,")
  InsertQuery.Add("                            :DATACREDITOBENEFICIARIO, :DATADEBITOBENEFICIARIO,")
  InsertQuery.Add("                            :DATACREDITORECEBEDOR, :DATADEBITORECEBEDOR,")
  InsertQuery.Add("                            :ACERTOLOTE, :TABCOBPAGACERTOSBENEF") ' O restante do insert será montado dinamicamente

  SQL.Clear
  SQL.Add("SELECT G.HANDLE, P.RECEBEDOR, G.PEG, G.EXECUTOR, G.LOCALEXECUCAO,  ")
  SQL.Add("       G.TABTIPOGUIA, G.BENEFICIARIO, G.DATAATENDIMENTO, G.TABREGIMEPGTO,")
  SQL.Add("       G.BANCO,G.AGENCIA, G.CONTACORRENTENUMERO,G.CONTACORRENTEDV,G.CONTACORRENTENOME,G.CONTACORRENTECPFCNPJ, G.AUTORIZACAO")
  SQL.Add("  FROM SAM_GUIA G, SAM_PEG P, SAM_COMPETPEG C")
  SQL.Add(" WHERE C.HANDLE = P.COMPETENCIA")
  SQL.Add("	  AND P.HANDLE = G.PEG")
  SQL.Add("	  AND G.SITUACAO = '4'")
  SQL.Add("	  AND G.HANDLE = :HANDLEGUIA")
  SQL.ParamByName("HANDLEGUIA").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    bsShowMessage("Guia sem Peg ou com Situação diferente de Faturada!", "E")
    CanContinue = False
    Exit Sub
  End If

  vReembolso = (SQL.FieldByName("TABREGIMEPGTO").AsInteger = 2)

  InsertQuery.ParamByName("HANDLE").AsInteger = NewHandle("SAM_PREACERTO_GUIA")
  InsertQuery.ParamByName("GUIA").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
  InsertQuery.ParamByName("DATAGERACAO").AsDateTime = ServerNow
  InsertQuery.ParamByName("USUARIOGERACAO").AsInteger = CurrentUser
  InsertQuery.ParamByName("ACERTOLOTE").AsInteger = RecordHandleOfTable("SAM_ACERTOLOTE")
  InsertQuery.ParamByName("ACERTOREALIZADO").AsString = "N"
  InsertQuery.ParamByName("PEG").AsInteger = SQL.FieldByName("PEG").AsInteger
  InsertQuery.ParamByName("DATAATENDIMENTO").AsDateTime = SQL.FieldByName("DATAATENDIMENTO").AsDateTime
  InsertQuery.ParamByName("DATACREDITOBENEFICIARIO").AsDateTime = ServerDate
  InsertQuery.ParamByName("DATADEBITOBENEFICIARIO").AsDateTime = ServerDate
  InsertQuery.ParamByName("DATACREDITORECEBEDOR").AsDateTime = ServerDate
  InsertQuery.ParamByName("DATADEBITORECEBEDOR").AsDateTime = ServerDate

  'If Not SQL.FieldByName("AUTORIZACAO").IsNull Then
  '  InsertQuery.Add(",:AUTORIZACAO")
  '  InsertQuery.ParamByName("AUTORIZACAO").AsInteger = SQL.FieldByName("AUTORIZACAO").AsInteger
  'Else
    InsertQuery.Add(",Null")
  'End If

  vHandleAux = SQL.FieldByName("EXECUTOR").AsInteger
  If vHandleAux <> 0 Then
    InsertQuery.Add(",:EXECUTOR")
    InsertQuery.ParamByName("EXECUTOR").AsInteger = vHandleAux
  Else
    InsertQuery.Add(",Null")
  End If

  vHandleAux = SQL.FieldByName("BENEFICIARIO").AsInteger
  If vHandleAux <> 0 Then
    InsertQuery.Add(",:BENEFICIARIO")
    InsertQuery.ParamByName("BENEFICIARIO").AsInteger = vHandleAux
  Else
    InsertQuery.Add(",Null")
  End If

   vHandleAux = SQL.FieldByName("LOCALEXECUCAO").AsInteger
  If vHandleAux <> 0 Then
    InsertQuery.Add(",:LOCALEXECUCAO")
    InsertQuery.ParamByName("LOCALEXECUCAO").AsInteger = vHandleAux
  Else
    InsertQuery.Add(",Null")
  End If

  If vReembolso Then
    vHandleAux = SQL.FieldByName("BANCO").AsInteger
    If vHandleAux <> 0 Then
      InsertQuery.Add(",:BANCO")
      InsertQuery.ParamByName("BANCO").AsInteger = SQL.FieldByName("BANCO").AsInteger
    Else
      InsertQuery.Add(",Null")
    End If
    vHandleAux = SQL.FieldByName("AGENCIA").AsInteger
    If vHandleAux <> 0 Then
      InsertQuery.Add(",:AGENCIA")
      InsertQuery.ParamByName("AGENCIA").AsInteger = SQL.FieldByName("AGENCIA").AsInteger
    Else
      InsertQuery.Add(",Null")
    End If
    If SQL.FieldByName("CONTACORRENTENUMERO").IsNull Then
      InsertQuery.Add(",Null")
    Else
      InsertQuery.Add(",:CONTACORRENTENUMERO")
      InsertQuery.ParamByName("CONTACORRENTENUMERO").AsString = SQL.FieldByName("CONTACORRENTENUMERO").AsString
    End If
    If Not SQL.FieldByName("CONTACORRENTEDV").IsNull Then
      InsertQuery.Add(",:CONTACORRENTEDV")
      InsertQuery.ParamByName("CONTACORRENTEDV").AsString = SQL.FieldByName("CONTACORRENTEDV").AsString
    Else
      InsertQuery.Add(",Null")
    End If
    If Not SQL.FieldByName("CONTACORRENTENOME").IsNull Then
      InsertQuery.Add(",:CONTACORRENTENOME")
      InsertQuery.ParamByName("CONTACORRENTENOME").AsString = SQL.FieldByName("CONTACORRENTENOME").AsString
    Else
      InsertQuery.Add(",Null")
    End If
    If Not SQL.FieldByName("CONTACORRENTECPFCNPJ").IsNull Then
      InsertQuery.Add(",:CONTACORRENTECPFCNPJ")
      InsertQuery.ParamByName("CONTACORRENTECPFCNPJ").AsString = SQL.FieldByName("CONTACORRENTECPFCNPJ").AsString
    Else
      InsertQuery.Add(",Null")
    End If

    If Not ((SQL.FieldByName("BANCO").IsNull) And (SQL.FieldByName("AGENCIA").IsNull) And (SQL.FieldByName("CONTACORRENTENUMERO").IsNull) _
        And (SQL.FieldByName("CONTACORRENTEDV").IsNull) And (SQL.FieldByName("CONTACORRENTENOME").IsNull) And (SQL.FieldByName("CONTACORRENTECPFCNPJ").IsNull)) Then
      InsertQuery.ParamByName("TABCOBPAGACERTOSBENEF").AsString = "3"
    Else
      InsertQuery.ParamByName("TABCOBPAGACERTOSBENEF").AsString = "1"
    End If

    InsertQuery.Add(",Null") 'Recebedor

  Else
    InsertQuery.Add(", Null, Null, Null, Null, Null, Null")
    If Not SQL.FieldByName("RECEBEDOR").IsNull Then
      InsertQuery.Add(",:RECEBEDOR")
      InsertQuery.ParamByName("RECEBEDOR").AsInteger = SQL.FieldByName("RECEBEDOR").AsInteger
    Else
      InsertQuery.Add(",Null")
    End If
    InsertQuery.ParamByName("TABCOBPAGACERTOSBENEF").AsString = "1"
  End If



  'SMS 44047 - Ricardo Vieira - 18/05/2005


  'FIM SMS 44047

  InsertQuery.Add(")")
'  bsShowMessage(InsertQuery.Text, "I")
  InsertQuery.ExecSQL

  Set InsertQuery = Nothing
  Set SQL = Nothing
End Sub
