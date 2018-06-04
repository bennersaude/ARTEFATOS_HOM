'HASH: FE1D78517EDBCF72C5691C1C3B0D1D30

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub BOTAOGERAREVENTOS_OnClick()

  Dim Duplica As Object
  Set Duplica = CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem, "SAM_CONTRATO_EVENTOS", "CONTRATO", RecordHandleOfTable("SAM_CONTRATO"), "Gerando eventos")
  Set Duplica = Nothing
  RefreshNodesWithTable "SAM_CONTRATO_EVENTOSEMAUTORIZ"

  Dim sql As Object
  Set sql = NewQuery

  If ( VisibleMode And NodeInternalCode = 10)  Then
	If Not InTransaction Then StartTransaction
    sql.Clear
    sql.Add("UPDATE SAM_CONTRATO_EVENTOS SET EXIGEAUTORIZACAO = 'S' WHERE CONTRATO = :CONTRATO And EXIGEAUTORIZACAO Is Null")
    sql.ParamByName("CONTRATO").AsInteger = RecordHandleOfTable("SAM_CONTRATO")
    sql.ExecSQL
    If InTransaction Then Commit
  Else
    If Not InTransaction Then StartTransaction
    sql.Clear
    sql.Add("UPDATE SAM_CONTRATO_EVENTOS SET EXIGEAUTORIZACAO = 'N' WHERE CONTRATO = :CONTRATO And EXIGEAUTORIZACAO Is Null")
    sql.ParamByName("CONTRATO").AsInteger = RecordHandleOfTable("SAM_CONTRATO")
    sql.ExecSQL
    If InTransaction Then Commit
  End If

  RefreshNodesWithTable("SAM_CONTRATO_EVENTOS")

  Set sql = Nothing
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)' só último nível
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOGERAREVENTOS"
			BOTAOGERAREVENTOS_OnClick
	End Select
End Sub
