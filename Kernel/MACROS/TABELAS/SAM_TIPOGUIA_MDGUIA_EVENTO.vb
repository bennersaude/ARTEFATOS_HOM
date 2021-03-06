﻿'HASH: 8E6D4E59988B45A91DF862F8D5CD3317
'#Uses "*bsShowMessage"

Dim vgORDEM As Integer

Public Sub RESEQUENCIAR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  Dim UPD As Object
  Set UPD = NewQuery

  SQL.Clear
  SQL.Add("SELECT * FROM SAM_TIPOGUIA_MDGUIA_EVENTO WHERE MODELOGUIA = :MODELOGUIA ORDER BY ORDEM")
  SQL.ParamByName("MODELOGUIA").Value = RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA")
  SQL.Active = True
  i = 10

  UPD.Clear
  If Not InTransaction Then StartTransaction
  UPD.Add("UPDATE SAM_TIPOGUIA_MDGUIA_EVENTO SET ORDEM = :ORDEM WHERE HANDLE = :HANDLE")

  i = 0

  While Not SQL.EOF
    UPD.ParamByName("ORDEM").Value = Format (Str(i + 10), "000")
    UPD.ParamByName("HANDLE").Value = SQL.FieldByName("HANDLE").Value
    UPD.ExecSQL
    i = UPD.ParamByName("ORDEM").Value
    SQL.Next
  Wend
  If InTransaction Then Commit
  RefreshNodesWithTable("SAM_TIPOGUIA_MDGUIA_EVENTO")

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
    OBRIGATORIEDADE.ReadOnly = False
  End If

  Set qTISS = Nothing


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT S1.DEPENDE, G.ORDEM, S2.ZCAMPO")
  SQL.Add("  FROM SIS_MODELOGUIA_CAMPOS S1, SIS_MODELOGUIA_CAMPOS S2, SAM_TIPOGUIA_MDGUIA_EVENTO G")
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
      bsShowMessage("Campo depende do " + SQL.FieldByName("ZCAMPO").AsString + "", "E")
      Set SQL = Nothing
      Exit Sub
    End If
  End If

  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_TIPOGUIA_MDGUIA_EVENTO WHERE SISCAMPO = :SISCAMPO AND HANDLE <> :HANDLE AND MODELOGUIA = :MODELOGUIA")
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

  'SMS 81587 - Débora Rebello - 16/05/2007
  If (Not VerificaInserirCodigoDespesas) Then
    MsgBox("Esse campo é exclusivo para modelos de guia do tipo Outras Despesas, do TISS.")
    CanContinue = False
    Exit Sub
  End If


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

  Set qCampo = Nothing

End Sub

Public Sub TABLE_NewRecord()
  i = vgORDEM
  CurrentQuery.FieldByName("ORDEM").Value = Format (Str(i + 10), "000")
  CurrentQuery.FieldByName("NOVALINHA").Value = "S"
End Sub

Public Function VerificaInserirCodigoDespesas() As Boolean
  'SMS 81587 - Débora Rebello - 16/05/2007
  Dim SQL As Object
  Dim qCampo As Object
  Set SQL = NewQuery
  Set qCampo = NewQuery

  qCampo.Active = False
  qCampo.Clear
  qCampo.Add("SELECT NOMECAMPO2")
  qCampo.Add("  FROM SIS_MODELOGUIA_CAMPOS")
  qCampo.Add(" WHERE HANDLE = :CAMPO")
  qCampo.ParamByName("CAMPO").AsInteger = CurrentQuery.FieldByName("SISCAMPO").AsInteger
  qCampo.Active = True

  If (qCampo.FieldByName("NOMECAMPO2").AsString = "EventoCodigoDespRealizadas") Then

	SQL.Active = False
  	SQL.Clear
  	SQL.Add("SELECT T.TIPOGUIATISS ")
	SQL.Add("  FROM SAM_TIPOGUIA T ")
  	SQL.Add("  JOIN SAM_TIPOGUIA_MDGUIA TM On (T.HANDLE = TM.TIPOGUIA) ")
  	SQL.Add(" WHERE TM.HANDLE = :HANDLE")
  	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_TIPOGUIA_MDGUIA")
  	SQL.Active = True

	'se o campo a ser inserido é EventoCodigoDespRealizadas, só pode inserir se o tipo de guia TISS não for outras despesas
	If (SQL.FieldByName("TIPOGUIATISS").AsString = "D") Then
	  VerificaInserirCodigoDespesas = True
  	Else
	  VerificaInserirCodigoDespesas = False
  	End If

  Else
    VerificaInserirCodigoDespesas = True
  End If

  Set SQL = Nothing
  Set qCampo = Nothing

End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "RESEQUENCIAR") Then
		RESEQUENCIAR_OnClick
	End If
End Sub
