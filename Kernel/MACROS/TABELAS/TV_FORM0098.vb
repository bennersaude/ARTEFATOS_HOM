'HASH: A7177492171031F6C2EA34AFD684C82C
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()

	If Not (WebMode) Then
		If bsShowMessage("Todas as guias do PEG serão glosadas. Deseja continuar?", "Q") = vbYes Then
			Dim vvSamPegDigit As Object
			Dim vbCriouGlosa As Boolean

			Set vvSamPegDigit = CreateBennerObject("SAMPEGDIGIT.Rotinas_GlosaTotal")

			vsMsg = vvSamPegDigit.GlosaTotalPeg(CurrentSystem, _
									   CLng(SessionVar("hPeg")), _
									   CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger, _
									   CurrentQuery.FieldByName("COMPLEMENTO").AsString, _
  									   vbCriouGlosa)
			If vsMsg <> "" Then
				bsShowMessage(vsMsg, "I")
			End If

    		Set vvSamPegDigit = Nothing

			Exit Sub
		End If
	End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If (WebMode) Then
		If bsShowMessage("Todas as guias do PEG serão glosadas. Deseja continuar?", "Q") = vbYes Then
			Dim vsMensagemErro As String

			Dim vvContainer As CSDContainer
			Set vvContainer = NewContainer

			vvContainer.AddFields("PEG:INTEGER;")
			vvContainer.AddFields("GUIA:INTEGER;")
			vvContainer.AddFields("MOTIVOGLOSA:INTEGER;")
			vvContainer.AddFields("COMPLEMENTO:STRING;")

			vvContainer.Insert
			vvContainer.Field("PEG").AsInteger = CLng(SessionVar("hPeg"))
			vvContainer.Field("GUIA").AsInteger = 0
			vvContainer.Field("MOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
			vvContainer.Field("COMPLEMENTO").AsString = CurrentQuery.FieldByName("COMPLEMENTO").AsString

			Dim Obj As Object
			Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                            		"SAMPEGDIGIT", _
	                                     "Rotinas_GlosaTotal", _
	                                     "Glosa Total do PEG n. " + SessionVar("hPeg"), _
	                                     0, _
	                                     "", _
	                                     "", _
	                                     "", _
	                                     "", _
	                                     "", _
	                                     True, _
	                                     vsMensagemErro, _
	                                     vvContainer)
			If viRetorno = 0 Then
				bsShowMessage("Processo enviado para execução no servidor!", "I")
			Else
				bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
			End If

			Set Obj = Nothing
			Set vvContainer = Nothing
		End If
	End If

End Sub
