'HASH: 63B4E6643B157CD6AA3D6922B98DA68D
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_AfterInsert()
	Dim query As BPesquisa
	Set query = NewQuery
	query.Clear

	If (SessionVar("ALTERARDATAS") = "atendimento") Then
		ROTULODATAATENDVALID.Text = "Alterar datas de atendimento de todos os eventos."
		query.Add(" SELECT MIN(DATAATENDIMENTO) DATA FROM SAM_AUTORIZ_EVENTOGERADO WHERE AUTORIZACAO = :HANDLE ")
	Else
		ROTULODATAATENDVALID.Text = "Alterar datas de validade de todos os eventos."
		NOVAHORA.ReadOnly = True
		NOVAHORA.Visible = False
		query.Add(" SELECT DATAVALIDADE DATA FROM SAM_AUTORIZ WHERE HANDLE = :HANDLE ")
	End If

	query.ParamByName("HANDLE").AsInteger = CInt(SessionVar("HANDLEAUTORIZACAO"))
	query.Active =True

	If (query.FieldByName("DATA").AsDateTime > 0) Then
		CurrentQuery.FieldByName("NOVADATA").AsDateTime = query.FieldByName("DATA").AsDateTime
	End If

	Set query = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If(CurrentQuery.FieldByName("NOVAHORA").IsNull And SessionVar("ALTERARDATAS") = "atendimento") Then
		bsShowMessage("Campo Hora é obrigatório!","I")
		CanContinue = False
		Exit Sub
	End If

	If(bsShowMessage("Este comando alterará a Data de " + SessionVar("ALTERARDATAS") + " de TODOS os eventos da autorização. Confirma?", "Q") = vbYes) Then

		Dim TvNovaDataBLL As CSBusinessComponent
		Set TvNovaDataBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.TvNovaDataEventosBLL, Benner.Saude.Atendimento.Business")
	    TvNovaDataBLL.AddParameter(pdtInteger, CInt(SessionVar("HANDLEAUTORIZACAO")))
	    TvNovaDataBLL.AddParameter(pdtString, SessionVar("ALTERARDATAS"))
   	    TvNovaDataBLL.AddParameter(pdtDateTime, CurrentQuery.FieldByName("NOVADATA").AsDateTime)
		Dim mensagem As String
		mensagem = TvNovaDataBLL.Execute("ValidarData")
		If Len(mensagem) > 0 Then
			bsShowMessage(mensagem,"E")
			CanContinue = False
		Else
		 	TvNovaDataBLL.AddParameter(pdtDateTime, CurrentQuery.FieldByName("NOVAHORA").AsDateTime)
			If (TvNovaDataBLL.Execute("ExecutarAlteracoes")) Then
				bsShowMessage("Alterações concluidas!","I")
			End If
		End If
		Set TvNovaDataBLL = Nothing
	Else
		bsShowMessage("Processo abortado!","I")
	End If
End Sub
