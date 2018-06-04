'HASH: DF4BBEEB972AC341FE87366FC4F31371


Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Duplica As Object
  Set Duplica = CreateBennerObject("SamDupEventos.Rotinas")
  Duplica.Duplicar(CurrentSystem, "SAM_PLANO_EVENTOS", "PLANO", RecordHandleOfTable("SAM_PLANO"), "Gerando eventos")
  Set Duplica = Nothing
  RefreshNodesWithTable "SAM_PLANO_EVENTOS"

  Dim sql As Object
  Set sql = NewQuery

  If (VisibleMode And NodeInternalCode = 10)  Then
    If Not InTransaction Then StartTransaction
    sql.Clear
    sql.Add("UPDATE SAM_PLANO_EVENTOS SET EXIGEAUTORIZACAO = 'S' WHERE PLANO = :PLANO And EXIGEAUTORIZACAO Is Null")
    sql.ParamByName("PLANO").AsInteger = RecordHandleOfTable("SAM_PLANO")
    sql.ExecSQL
    If InTransaction Then Commit
  Else
    If Not InTransaction Then StartTransaction
    sql.Clear
    sql.Add("UPDATE SAM_PLANO_EVENTOS SET EXIGEAUTORIZACAO = 'N' WHERE PLANO = :PLANO And EXIGEAUTORIZACAO Is Null")
    sql.ParamByName("PLANO").AsInteger = RecordHandleOfTable("SAM_PLANO")
    sql.ExecSQL
    If InTransaction Then Commit

  End If

  RefreshNodesWithTable("SAM_PLANO_EVENTOS")

  Set sql = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOGERAREVENTOS" Then
		BOTAOGERAREVENTOS_OnClick
	End If
End Sub
