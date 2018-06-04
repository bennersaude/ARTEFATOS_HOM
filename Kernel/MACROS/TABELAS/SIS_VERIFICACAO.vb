'HASH: E3FBC304335DA95BC53120F1851C28B4
'macro: sis_verificação
'Juliano -02/04/2001 -processo 1

'Última alteração: 23/04/2002
'Milton -SMS 8698

'#Uses "*bsShowMessage"

Dim Obj As Object
Option Explicit

Public Sub Executar
	If VisibleMode Then
		Set Obj = CreateBennerObject("BSInterface0019.Rotinas")
			Obj.ExecutarVerificacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	Else
		Dim vsMensagemErro As String
		Dim viRetorno As Long

		Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                   						"SISVERIFICACAO", _
		   	                     		"Rotinas", _
		       	                 		"Rotina de verificação - Verificação: " + _
		           	             		CStr(CurrentQuery.FieldByName("HANDLE").AsInteger) + _
			           	         		" Descrição: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
		                       			CurrentQuery.FieldByName("HANDLE").AsInteger, _
		                   	     		"SIS_VERIFICACAO", _
		                      	 		"SITUACAO", _
		                       			"", _
		                       			"", _
		                       			"P", _
		                       			True, _
		                       			vsMensagemErro, _
		                       			Null)
		If viRetorno = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
		Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		End If
	End If
End Sub


Public Sub BOTAOPROCESSAR_OnClick()

	If CurrentQuery.FieldByName("CODIGO").AsInteger = 01 Then
		If bsShowMessage("Confirma verificação de todas as matrículas?", "Q") = vbYes Then
			Executar
		Else
			Exit Sub
		End If
	ElseIf CurrentQuery.FieldByName("CODIGO").AsInteger = 08 Then
		If WebMode Then
			bsShowMessage("Verificação não pode ser executada em modo WEB", "I")
			Exit Sub
		Else
			If bsShowMessage("Confirma verificação de todos Prestadores sem Conta Financeira ?", "Q") = vbYes Then
				Executar
			End If
		End If
	ElseIf CurrentQuery.FieldByName("CODIGO").AsInteger = 18 Then
		If bsShowMessage("Confirma verificação de todas os eventos?", "Q") = vbYes Then
			Executar
		End If
	ElseIf (WebMode) And (CurrentQuery.FieldByName("CODIGO").AsInteger = 15) Then
		bsShowMessage("Verificação não pode ser executada em modo WEB", "I")
		Exit Sub
	ElseIf (WebMode) And (CurrentQuery.FieldByName("CODIGO").AsInteger = 20) Then
		bsShowMessage("Verificação não pode ser executada em modo WEB", "I")
		Exit Sub
	ElseIf CurrentQuery.FieldByName("CODIGO").AsInteger = 26 Then
		If bsShowMessage("Ao processar está verificação, será alterado as descrições de todos os eventos para o padrão da TUSS!" + Chr(13) + "Deseja continuar?", "Q") = vbYes Then
		    Executar
		Else
			Exit Sub
		End If
	ElseIf (WebMode) And (CurrentQuery.FieldByName("CODIGO").AsInteger = 31) Then
		bsShowMessage("Verificação não pode ser executada em modo WEB", "I")
		Exit Sub
	ElseIf CurrentQuery.FieldByName("CODIGO").AsInteger = 27 Then
	  Dim qParametro As Object 'Coelho SMS: 121102 criação do parametro geral
      Set qParametro = NewQuery

      qParametro.Clear
      qParametro.Add("SELECT REPLICAEQUIVALENTES FROM SAM_PARAMETROSPROCCONTAS")
      qParametro.Active = True
      If qParametro.FieldByName("REPLICAEQUIVALENTES").AsString = "S" Then
         Executar
      Else
         Set qParametro = Nothing
         bsShowMessage("Processo abortado. Parâmetro geral ""Replica eventos equivalentes"" está desmarcado!", "I")
		 Exit Sub
      End If
	Else
		Executar
	End If

	RefreshNodesWithTable("SIS_VERIFICACAO")

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "BOTAOPROCESSAR") Then
		BOTAOPROCESSAR_OnClick
	End If
End Sub
