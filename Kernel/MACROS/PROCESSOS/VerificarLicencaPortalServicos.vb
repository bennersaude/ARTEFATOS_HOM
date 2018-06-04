'HASH: B799B131EB8096DCBA28C7CE208B4290
Public Sub Main

	Dim resultado As String

	resultado = "S"

	On Error GoTo Erro

		Dim serverExec As CSServerExec

		Set serverExec = NewServerExec

		serverExec.Description = "Processar verificação de Licença do Portal de Serviços"
		serverExec.DllClassName = "Benner.Saude.Web.PortalServicos.ManagerInstaladorFuncionalidades.ProcessarVerificarLicenca"

		serverExec.Execute
		serverExec.Wait

		If (VisibleMode) Then
			MsgBox("Processamento de verificação da licença enviado para o servidor. Aguarde alguns minutos antes de acessar o Portal Serviços novamente.")
		Else
			InfoDescription = "Processamento de verificação da licença enviado para o servidor. Aguarde alguns minutos antes de acessar o Portal Serviços novamente."
		End If

		GoTo Finaliza
	Erro:
		If (VisibleMode) Then
			MsgBox(Err.Description)
		Else
			InfoDescription = Err.Description
		End If

		resultado = "N"

	Finaliza:
		If (serverExec.Status = esError) Then
			If (VisibleMode) Then
				MsgBox(serverExec.ErrorMessage)
			Else
				InfoDescription = serverExec.ErrorMessage
			End If

			resultado = "N"
		End If


	Set serverExec = Nothing

	ServiceResult = resultado

End Sub

