'HASH: 0965937C257A32DEA9DD438FAB196776
'Macro: SAM_PRECOPRESTADOR_CASCATA_QTD
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Dim SQL
	Set SQL = NewQuery

	SQL.Add("SELECT DATAFINAL FROM SAM_PRECOPRESTADOR_CASCATA  WHERE HANDLE = :HPRESTCOCASCATA")

	SQL.ParamByName("HPRESTCOCASCATA").Value = RecordHandleOfTable("SAM_PRECOPRESTADOR_CASCATA")
	SQL.Active = True

	If Not SQL.FieldByName("DATAFINAL").IsNull Then
		bsShowMessage("Item da Cascata finalizado não permite manutenções", "E")
		CurrentQuery.Cancel
		RefreshNodesWithTable("SAM_PRECOPRESTADOR_CASCATA_QTD")
	End If

	SQL.Active = False

	Set SQL = Nothing
End Sub
