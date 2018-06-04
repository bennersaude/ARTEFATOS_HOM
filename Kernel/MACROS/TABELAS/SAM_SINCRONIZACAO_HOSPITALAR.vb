'HASH: 2A96DA3D550CA7D41C00D8114EAED831
'#Uses "*bsShowMessage"

Public Sub BOTAOINCLUIRNAFILA_AfterOnClick()
	RefreshNodesWithTable("SAM_SINCRONIZACAO_HOSPITALAR")
End Sub

Public Sub BOTAOINCLUIRNAFILA_BeforeOnClick(CanContinue As Boolean)
	If (bsShowMessage("Confirma a inclusão do registro na fila de sincronização?", "Q") = vbNo) Then
		CanContinue = False
		Exit Sub
	End If
End Sub
