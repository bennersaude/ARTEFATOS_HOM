'HASH: D3E4E644596F2E8259D1DC89247C0505
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	excluirProblemas(CurrentQuery.FieldByName("HANDLE").AsInteger)
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "SUSPENDER") Then
		InfoDescription = suspender
	End If
End Sub

Public Function suspender As String
	On Error GoTo erro
	CurrentQuery.Edit
	CurrentQuery.FieldByName("DATAFINAL").AsDateTime = CurrentVirtualQuery.FieldByName("DATAFINAL").AsDateTime
	CurrentQuery.Post
	suspender = ""
	Exit Function
erro:
	suspender = Err.Description
End Function


Public Sub excluirProblemas(handle As Long)
	Dim sql As BPesquisa
	Set sql=NewQuery()
	sql.Add("DELETE FROM CLI_MEDICAMENTOUSOCONTINUO_CID WHERE MEDICAMENTOUSOCONTINUO=:M")
	sql.ParamByName("M").AsInteger = handle
	sql.ExecSQL
	Set sql=Nothing
End Sub
