'HASH: FF0EC879D8EC914D4369FC7ABD8A03E6
 

Public Sub BOTAOCANCELAR_OnClick()

	'só permite cancelar na situacao Processada
	If CurrentQuery.FieldByName("SITUACAO").AsInteger <> 4 Then
		MsgBox("Só é permitido cancelar na situação Aguardando importação.", vbInformation, "Aviso")
		Exit Sub
	End If

	If MsgBox("Deseja realmente cancelar a exportação?", vbYesNo) = vbYes Then

		On Error GoTo erro:

			StartTransaction

			Dim strOcorrencia As String

			Dim sqlBuscarUsuario As BPesquisa
			Set sqlBuscarUsuario = NewQuery

			sqlBuscarUsuario.Add("SELECT APELIDO FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HANDLE")
			sqlBuscarUsuario.ParamByName("HANDLE").Value = CurrentUser
			sqlBuscarUsuario.Active = True

			strOcorrencia = "Cancelada por: "+sqlBuscarUsuario.FieldByName("APELIDO").AsString

			strOcorrencia = strOcorrencia + " em " + FormatDateTime2("DD/MM/YYYY HH:MM:SS", ServerNow) + Chr(13)

			Set sqlBuscarUsuario = Nothing

			Dim sqlUpdate As BPesquisa
			Set sqlUpdate = NewQuery

			If Not InTransaction Then
				StartTransaction
			End If

			sqlUpdate.Add("DELETE FROM GRS_EXPORTACAO_XML WHERE EXPORTACAOPREVINNE = :ROTINA")
			sqlUpdate.ParamByName("ROTINA").Value = CurrentQuery.FieldByName("HANDLE").Value
			sqlUpdate.ExecSQL

			sqlUpdate.Clear
			sqlUpdate.Add("UPDATE GRS_EXPORTACAO_PREVINNE SET SITUACAO = :SITUACAO, OCORRENCIAS = :OCORRENCIA, ")
			sqlUpdate.Add(" USUARIOCANCELAMENTO = :USUARIOCANCELAMENTO, DATAHORACANCELAMENTO = :DATAHORACANCELAMENTO ")
			sqlUpdate.Add(" WHERE HANDLE = :HANDLE")
			sqlUpdate.ParamByName("SITUACAO").Value = 5
			sqlUpdate.ParamByName("OCORRENCIA").Value = CurrentQuery.FieldByName("OCORRENCIAS").AsString + Chr(13) + "-----------------------------------" + Chr(13) + strOcorrencia
			sqlUpdate.ParamByName("USUARIOCANCELAMENTO").AsInteger = CurrentUser
			sqlUpdate.ParamByName("DATAHORACANCELAMENTO").AsDateTime = ServerNow()
			sqlUpdate.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
			sqlUpdate.ExecSQL

			Set sqlUpdate = Nothing

			WriteAudit("A", HandleOfTable("GRS_EXPORTACAO_PREVINNE"), CLng(CurrentQuery.FieldByName("HANDLE").Value), "Cancelada a rotina de exportação")

			Commit

			Exit Sub

		erro:

			Rollback

			MsgBox("Erro ao cancelar exportação. "+Err.Description, vbCritical, "Aviso")

	End If

End Sub

Public Sub BOTAOEXPORTAR_OnClick()
	Dim vi_Situacao As Integer

	vi_Situacao = CurrentQuery.FieldByName("SITUACAO").AsInteger

	'somente será realizada a exportacao quando a rotina estiver aguardando exportação
	If ( vi_Situacao <> 4 ) Then
		MsgBox("Só é permitido exportar na situação 'Aguardando importação'.", vbInformation, "Aviso")
		Exit Sub
	End If

	Dim status As String

	On Error GoTo erroQuery
		Dim dll As Object
		Set dll = CreateBennerObject("EnviarExportacaoBeneficiarioXml.ExecutarExportacao")
		dll.Exportar( CurrentSystem, status )
		Set dll = Nothing

 		GoTo fim

	erroQuery:
			MsgBox(status + " - Description: " + Err.Description, vbInformation, "Erro no processo")

			GoTo fim
	fim:
			If Not status = "" Then
				MsgBox(status, vbInformation, "Aviso")
			End If
End Sub

Public Sub BOTAOGERAR_OnClick()

	'verifica se o registro esta sendo incluido ou alterado
	If CurrentQuery.InInsertion Or CurrentQuery.InEdition Then
		MsgBox("Confirme a rotina antes de exportar.", vbInformation, "Aviso")
		Exit Sub
	End If

	Dim vi_Situacao As Integer

	vi_Situacao = CurrentQuery.FieldByName("SITUACAO").AsInteger

	'somente será realizada a exportacao quando a rotina estiver pendente, aguardando exportação ou cancelada
	If ( vi_Situacao <> 1 ) And ( vi_Situacao <> 5 ) Then
		MsgBox("Só é permitido exportar nas situações 'Pendente', 'Aguardando importação' e 'Cancelada'.", vbInformation, "Aviso")
		Exit Sub
	End If

	Dim sql As BPesquisa

	Set sql = NewQuery

	sql.Clear
	sql.Add("SELECT COUNT(*) PROCESSANDO ")
	sql.Add(" FROM GRS_EXPORTACAO_PREVINNE ")
	sql.Add(" WHERE SITUACAO = 2 ")
	sql.Active = True

	'nao exportar se houver outra exportacao em andamento
	If (sql.FieldByName("PROCESSANDO").AsInteger > 0) Then
		MsgBox("Existe outra exportacao em andamento.", vbInformation, "Aviso")
		Exit Sub
	End If

	Dim sqlUpdate As BPesquisa
	Set sqlUpdate = NewQuery

	If Not InTransaction Then
		StartTransaction
	End If

	sqlUpdate.Clear
	sqlUpdate.Add("UPDATE GRS_EXPORTACAO_PREVINNE SET ")
	sqlUpdate.Add(" USUARIOEXPORTACAO = :USUARIOEXPORTACAO, DATAHORAEXPORTACAO = :DATAHORAEXPORTACAO ")
	sqlUpdate.Add(" WHERE HANDLE = :HANDLE")
	sqlUpdate.ParamByName("USUARIOEXPORTACAO").AsInteger = CurrentUser
	sqlUpdate.ParamByName("DATAHORAEXPORTACAO").AsDateTime = ServerNow()
	sqlUpdate.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
	sqlUpdate.ExecSQL

	Set sqlUpdate = Nothing

	Commit

	Dim status As String

	On Error GoTo erroQuery
		Dim dll As Object
		Set dll = CreateBennerObject("ExportacaoBeneficiarioXML.ExportacaoBeneficiario")
		dll.Exportar( CurrentSystem, CurrentQuery.FieldByName("DATAINIFILTRO").AsDateTime, CurrentQuery.FieldByName("DATAFIMFILTRO").AsDateTime, status )
		Set dll = Nothing

 		GoTo fim

	erroQuery:
			MsgBox(status + " - Description: " + Err.Description, vbInformation, "Erro no processo")

			GoTo fim
	fim:
			If Not status = "" Then
				MsgBox(status, vbInformation, "Aviso")
			End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	'só permite apagar uma rotina na situacao Pendente ou Cancelada
	If (CurrentQuery.FieldByName("SITUACAO").AsInteger <> 1) And (CurrentQuery.FieldByName("SITUACAO").AsInteger <> 5) Then
		MsgBox("Só é permitido apagar exportação nas situações 'Pendente' e 'Cancelada'.", vbInformation, "Aviso")
		CanContinue = False
	End If

End Sub
