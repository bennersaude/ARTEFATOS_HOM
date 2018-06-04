'HASH: 6A09598B45515371A8E534C0ABC35C43
 Option Explicit

Public Sub TABLE_AfterDelete()
		RefreshNodesWithTable "SAM_EXCEPCIONALIDADE"
End Sub

Public Sub TABLE_NewRecord()
	CurrentEntity.TransitoryVars("HANDLEPRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
	CurrentEntity.TransitoryVars("HANDLETIPOPRESTADOR").AsInteger = RecordHandleOfTable("SAM_TIPOPRESTADOR")
	CurrentEntity.TransitoryVars("HANDLEMOTIVOGLOSA").AsInteger = RecordHandleOfTable("SAM_MOTIVOGLOSA")
End Sub

Public Sub TABLE_OnDeleteBtnClick(CanContinue As Boolean)
	PreencheTransitorVars
End Sub

Public Sub TABLE_OnSaveBtnClick(CanContinue As Boolean)
	PreencheTransitorVars
End Sub
Public Sub PreencheTransitorVars
	CurrentEntity.TransitoryVars("HANDLEPRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
	CurrentEntity.TransitoryVars("HANDLETIPOPRESTADOR").AsInteger = RecordHandleOfTable("SAM_TIPOPRESTADOR")
	CurrentEntity.TransitoryVars("HANDLEMOTIVOGLOSA").AsInteger = RecordHandleOfTable("SAM_MOTIVOGLOSA")
	CurrentEntity.TransitoryVars("HANDLEEXCEPCIONALIDADE").AsInteger = RecordHandleOfTable("SAM_EXCEPCIONALIDADE")
End Sub
