'HASH: 8CB4F82F288A8C266053C3923CD2357F
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "TIPOARQUIVOPDF" Then
		UserVar("TIPOARQUIVO") = ".pdf"
		InfoDescription = "O formato das novas solicitações de relatórios foi alterado para ser gerado em PDF."
	ElseIf CommandID = "TIPOARQUIVOTXT" Then
		UserVar("TIPOARQUIVO") = ".txt"
		InfoDescription = "O formato das novas solicitações de relatórios foi alterado para ser gerado em TXT."
	ElseIf CommandID = "TIPOARQUIVODOC" Then
		UserVar("TIPOARQUIVO") = ".docx"
		InfoDescription = "O formato das novas solicitações de relatórios foi alterado para ser gerado em DOC."
	ElseIf CommandID = "TIPOARQUIVOXLS" Then
		UserVar("TIPOARQUIVO") = ".xlsx"
		InfoDescription = "O formato das novas solicitações de relatórios foi alterado para ser gerado em XLS."
	ElseIf CommandID = "TIPOARQUIVOCSV" Then
		UserVar("TIPOARQUIVO") = ".csv"
		InfoDescription = "O formato das novas solicitações de relatórios foi alterado para ser gerado em CSV."
	End If
End Sub
