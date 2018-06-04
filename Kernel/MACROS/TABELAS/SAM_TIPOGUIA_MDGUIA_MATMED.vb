'HASH: C6EA8C5F4F2A84624E982FB3CD781C16
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
