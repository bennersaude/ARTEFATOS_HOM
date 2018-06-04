'HASH: 9E63A99BDF499F385FA3D442DFA2E211
'#Uses "*bsShowMessage"

Dim vgORDEM As Integer


Public Sub RESEQUENCIAR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  Dim UPD As Object
  Set UPD = NewQuery

  SQL.Clear
  SQL.Add("SELECT * FROM SAM_TIPOGUIA_MDGUIA_GUIA WHERE MODELOGUIA = :MODELOGUIA ORDER BY ORDEM")
  SQL.ParamByName("MODELOGUIA").Value = RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA")
  SQL.Active = True
  i = 10

  UPD.Clear
  If Not InTransaction Then StartTransaction
  UPD.Add("UPDATE SAM_TIPOGUIA_MDGUIA_GUIA SET ORDEM = :ORDEM WHERE HANDLE = :HANDLE")

  i = 0

  While Not SQL.EOF
    UPD.ParamByName("ORDEM").Value = Format (Str(i + 10), "000")
    UPD.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
    UPD.ExecSQL
    i = UPD.ParamByName("ORDEM").Value
    SQL.Next
  Wend
  If InTransaction Then Commit
  RefreshNodesWithTable("SAM_TIPOGUIA_MDGUIA_GUIA")

End Sub

Public Sub SISCAMPO_OnExit()
  If CurrentQuery.State <> 1 Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Add("SELECT LARGURAPADRAO, LEGENDA FROM SIS_MODELOGUIA_CAMPOS WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SISCAMPO").AsInteger
    SQL.Active = True

    CurrentQuery.FieldByName("LEGENDA").Value = SQL.FieldByName("LEGENDA").Value
    'UCase(Mid(SISCAMPO.Text,1,1)) + LCase(Mid(SISCAMPO.Text,2))
    CurrentQuery.FieldByName("LARGURA").Value = SQL.FieldByName("LARGURAPADRAO").Value

    Set SQL = Nothing

  End If
End Sub


Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("ORDEM").IsNull Then
    vgORDEM = CurrentQuery.FieldByName("ORDEM").AsInteger
  End If

  'sms 77454
  Dim qTISS As Object
  Set qTISS = NewQuery
  qTISS.Active = False
  qTISS.Add("SELECT G.TIPOGUIATISS")
  qTISS.Add("  FROM SAM_TIPOGUIA_MDGUIA M, SAM_TIPOGUIA G")
  qTISS.Add(" WHERE M.TIPOGUIA = G.Handle And M.Handle = :MODELOGUIA")
  'qTISS.ParamByName("MODELOGUIA").AsInteger = CurrentQuery.FieldByName("MODELOGUIA").AsInteger - SMS 81195 - Débora Rebello - 07/05/2007
  qTISS.ParamByName("MODELOGUIA").AsInteger = RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA") 'SMS 81195 - Débora Rebello - 07/05/2007
  qTISS.Active = True

  If (qTISS.FieldByName("TIPOGUIATISS").AsString ="N") Or (qTISS.FieldByName("TIPOGUIATISS").IsNull) Then
    OBRIGATORIEDADE.ReadOnly = True
  Else
    OBRIGATORIEDADE.ReadOnly = False 'SMS 81195 - Débora Rebello - 07/05/2007
  End If

  Set qTISS = Nothing


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaSeJahTemCampoData As Object
  Set qVerificaSeJahTemCampoData = NewQuery

  qVerificaSeJahTemCampoData.Clear
  qVerificaSeJahTemCampoData.Add("SELECT HANDLE FROM SIS_MODELOGUIA_CAMPOS WHERE HANDLE = :CAMPO AND EVENTOGUIA = 'G' AND NOMECAMPO = 'DataAtend'")
  qVerificaSeJahTemCampoData.ParamByName("CAMPO").AsInteger = CurrentQuery.FieldByName("SISCAMPO").AsInteger
  qVerificaSeJahTemCampoData.Active = True
  If qVerificaSeJahTemCampoData.FieldByName("HANDLE").AsInteger > 0 Then 'o campo a ser incluído é o campo de dataatend da guia
    qVerificaSeJahTemCampoData.Clear
    qVerificaSeJahTemCampoData.Add("SELECT MDGG.HANDLE HANDLE")
    qVerificaSeJahTemCampoData.Add("  FROM SAM_TIPOGUIA_MDGUIA_GUIA MDGG,")
    qVerificaSeJahTemCampoData.Add("       SIS_MODELOGUIA_CAMPOS    MDC  ")
    qVerificaSeJahTemCampoData.Add(" WHERE MDGG.SISCAMPO = MDC.HANDLE")
    qVerificaSeJahTemCampoData.Add("   And MDC.EVENTOGUIA = 'G'")
    qVerificaSeJahTemCampoData.Add("   And MDC.NOMECAMPO = 'DataAtend'")
    qVerificaSeJahTemCampoData.Add("   And MDGG.MODELOGUIA = :MODELOGUIA")
    qVerificaSeJahTemCampoData.Add("   And MDGG.HANDLE <> :HANDLE")
    qVerificaSeJahTemCampoData.ParamByName("MODELOGUIA").AsInteger = CurrentQuery.FieldByName("MODELOGUIA").AsInteger
    qVerificaSeJahTemCampoData.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qVerificaSeJahTemCampoData.Active = True
    If qVerificaSeJahTemCampoData.FieldByName("HANDLE").AsInteger > 0 Then
      bsShowMessage("Já possui o Campo DataAtend neste modelo de guia", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Set qVerificaSeJahTemCampoData = Nothing
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT S1.DEPENDE, G.ORDEM, S2.ZCAMPO")
  SQL.Add("  FROM SIS_MODELOGUIA_CAMPOS S1, SIS_MODELOGUIA_CAMPOS S2, SAM_TIPOGUIA_MDGUIA_GUIA G")
  SQL.Add(" WHERE S1.DEPENDE = S2.HANDLE")
  SQL.Add("   AND G.SISCAMPO = S2.HANDLE")
  SQL.Add("   AND S1.HANDLE  = :HANDLE")
  SQL.Add("   AND G.MODELOGUIA = :MODELOGUIA")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SISCAMPO").AsInteger
  SQL.ParamByName("MODELOGUIA").Value = RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA")
  SQL.Active = True
  If (Not SQL.EOF) Then
    If (Not SQL.FieldByName("DEPENDE").IsNull) And _
        (CurrentQuery.FieldByName("ORDEM").AsInteger <= SQL.FieldByName("ORDEM").AsInteger) Then
      CanContinue = False
      bsShowMessage("Campo depende do " + SQL.FieldByName("ZCAMPO").AsString + "","E")
      Set SQL = Nothing
      Exit Sub
    End If
  End If

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA_GUIA WHERE SISCAMPO = :SISCAMPO And HANDLE <> :HANDLE And MODELOGUIA = :MODELOGUIA")
  SQL.ParamByName("SISCAMPO").Value = CurrentQuery.FieldByName("SISCAMPO").AsInteger
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("MODELOGUIA").Value = RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA")
  SQL.Active = True

  If (Not SQL.EOF) Then
    CanContinue = False
    bsShowMessage("Campo já cadastrado.", "E")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  'sms 77454
  Dim vModeloTISSOk As String
  Dim vCampoEhObrigatorio As String
  Dim qCampo As Object
  Set qCampo = NewQuery

  qCampo.Active = False
  qCampo.Add("SELECT NOMECAMPO2")
  qCampo.Add("  FROM SIS_MODELOGUIA_CAMPOS")
  qCampo.Add(" WHERE HANDLE = :CAMPO")
  qCampo.ParamByName("CAMPO").AsInteger = CurrentQuery.FieldByName("SISCAMPO").AsInteger
  qCampo.Active = True

  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("BsPro006.Geral")
  INTERFACE.VerificaModeloTISS(CurrentSystem, CurrentQuery.FieldByName("MODELOGUIA").AsInteger, qCampo.FieldByName("NOMECAMPO2").AsString, _
                               vModeloTISSOk, vCampoEhObrigatorio)
  Set INTERFACE = Nothing

  If (vCampoEhObrigatorio= "S") And (CurrentQuery.FieldByName("OBRIGATORIEDADE").AsString <> "1") Then
    bsShowMessage("Campo obrigatório conforme tipo de guia TISS. Obrigatoriedade deve estar marcado 'Obrigatório'.", "E")
    CanContinue = False
    Exit Sub
  End If



End Sub

Public Sub TABLE_NewRecord()
  i = vgORDEM
  CurrentQuery.FieldByName("ORDEM").Value = Format (Str(i + 10), "000")
  CurrentQuery.FieldByName("NOVALINHA").Value = "S"
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "RESEQUENCIAR") Then
		RESEQUENCIAR_OnClick
	End If
End Sub
