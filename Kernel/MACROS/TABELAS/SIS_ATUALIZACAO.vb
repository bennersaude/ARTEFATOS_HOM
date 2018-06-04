'HASH: DB41432CDE2BD8CFB32285B1285452C3
'MACRO: sis_atualização
'#Uses "*bsShowMessage"
Option Explicit

Public Sub BOTAOPROCESSAR_OnClick()
Dim Obj As Object

'SMS 90292 - Ricardo Rocha - Adequacao para WEB

	If ((CurrentQuery.FieldByName("SITUACAO").AsString = "1") And _
		(CurrentQuery.FieldByName("PROCESSADOUSUARIO").AsInteger > 0 ) And _
		(CurrentQuery.FieldByName("PROCESSADODATA").AsDateTime > 0)) Then

		Dim QueryAtualiza As Object
		Set QueryAtualiza = NewQuery

		QueryAtualiza.Add("UPDATE SIS_ATUALIZACAO  ")
		QueryAtualiza.Add(" SET SITUACAO = '5'     ")
		QueryAtualiza.Add(" WHERE HANDLE = :HANDLE ")

		QueryAtualiza.ParamByName("HANDLE").AsInteger  = CurrentQuery.FieldByName("HANDLE").AsInteger

		QueryAtualiza.ExecSQL

		Set QueryAtualiza = Nothing
	End If

	If CurrentQuery.FieldByName("CODIGO").AsInteger < 60 Then
		bsShowMessage("Processo desta atualização ainda não implementado", "I")
		Exit Sub
	ElseIf CurrentQuery.FieldByName("CODIGO").AsInteger = 1361 Then
		Dim vsMsg As String

		vsMsg = "Serão marcados como nulo TODOS os campos que dependem " + Chr(13) _
		+ "dos registros da tabela TIS_TABELAPRECO. São eles: " + Chr(13) + Chr(13) _
		+ " - SAM_AUTORIZ_EVENTOSOLICIT.CODIGOTABEL" + Chr(13) _
		+ " - SAM_GUIA_EVENTOS.CODIGOTABELA" + Chr(13) _
		+ " - SAM_GUIA_EVENTOS.TABELAPRECO" + Chr(13) _
		+ " - SAM_PARAMETROSWEB.CODIGOTABELA" + Chr(13) _
		+ " - WEB_AUTORIZ.CODIGOTABELA" + Chr(13) _
		+ " - WEB_GUIA.TABELAPRECO" + Chr(13) _
		+ " - WEB_GUIA_EVENTOS.TABELA" + Chr(13) + Chr(13) _
		+ "Deseja continuar?"

		If bsShowMessage(vsMsg, "Q") = vbYes Then
			If VisibleMode Then
				Set Obj = CreateBennerObject("BSInterface0019.Rotinas")
				Obj.ExecutarAtualizacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
			Else
				Dim vsMensagemErro As String
				Dim viRetorno As Long

				Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
				viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                         						"SISATUALIZACAO", _
				                         		"Rotinas", _
				                         		"Rotina de atualização de sistema - Atualização: " + _
				                         		CStr(CurrentQuery.FieldByName("HANDLE").AsInteger) + _
				                         		" Descrição: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
				                         		CurrentQuery.FieldByName("HANDLE").AsInteger, _
				                         		"SIS_ATUALIZACAO", _
				                         		"SITUACAO", _
				                         		"", _
				                         		"", _
				                         		"P", _
				                         		False, _
				                         		vsMensagemErro, _
				                         		Null)

				If viRetorno = 0 Then
					bsShowMessage("Processo enviado para execução no servidor!", "I")
				Else
					bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
				End If
			End If
			Set Obj = Nothing
		End If
	ElseIf (WebMode) And ((CurrentQuery.FieldByName("CODIGO").AsInteger = 466) Or (CurrentQuery.FieldByName("CODIGO").AsInteger = 127) _
					Or (CurrentQuery.FieldByName("CODIGO").AsInteger = 129) Or (CurrentQuery.FieldByName("CODIGO").AsInteger = 164)) Then
		bsShowMessage("Atualização não pode ser executada em modo WEB", "I")
		Exit Sub
	ElseIf CurrentQuery.FieldByName("CODIGO").AsInteger = 1398 Then

	  If Not CurrentQuery.FieldByName("PROCESSADOUSUARIO").IsNull Then
         bsShowMessage("Atualização já processada", "I")
         Exit Sub
      End If

      Set Obj = CreateBennerObject("SisAtualizacao.Rotinas")
      Obj.Exec(CurrentSystem, CurrentQuery.FieldByName("CODIGO").AsInteger, CurrentQuery.FieldByName("DESCRICAO").AsString)

      Set Obj = Nothing

      CurrentQuery.Active = False
      CurrentQuery.Active = True
	Else
		If VisibleMode Then
			Set Obj = CreateBennerObject("BSInterface0019.Rotinas")
			Obj.ExecutarAtualizacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
		Else
			Dim vsMsgErro As String
			Dim viRet As Long

			Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRet = Obj.ExecucaoImediata(CurrentSystem, _
			                       		"SISATUALIZACAO", _
			                       		"Rotinas", _
			                       		"Rotina de atualização de sistema - Atualização: " + _
			                       		CStr(CurrentQuery.FieldByName("HANDLE").AsInteger) + _
			                       		" Descrição: " + CurrentQuery.FieldByName("DESCRICAO").AsString, _
			                       		CurrentQuery.FieldByName("HANDLE").AsInteger, _
			                       		"SIS_ATUALIZACAO", _
			                       		"SITUACAO", _
			                       		"", _
			                       		"", _
			                       		"P", _
			                       		False, _
			                       		vsMsgErro, _
			                       		Null)

			If viRet = 0 Then
				bsShowMessage("Processo enviado para execução no servidor!", "I")
			Else
				bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMsgErro, "I")
			End If
			Set Obj = Nothing
		End If
	End If

RefreshNodesWithTable("SIS_ATUALIZACAO")

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	End If
End Sub
