'HASH: C77C7C0959065C22719F1AA69CEBB3DB
'#Uses "*bsShowMessage"

Public Sub TABLE_OnDeleteBtnClick(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("SITUACAO").AsString <> "1") Then
		bsShowMessage("Não é possível Excluir uma rotina que não esteja aberta", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_OnEditBtnClick(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("SITUACAO").AsString <> "1") Then
		bsShowMessage("Não é possível Editar uma rotina que não esteja aberta", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_OnSaveBtnClick(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("SITUACAO").AsString <> "1") Then
		bsShowMessage("Não é possível alterar uma rotina que não esteja aberta", "E")
		CanContinue = False
	End If
End Sub

Public Sub BOTAOPROCESSAR_AfterOnClick()
	RefreshNodesWithTable("SAM_FECHAMENTOPRECO")
End Sub
