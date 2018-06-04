'HASH: 288773DCFFE9DF17A654674ED8692B50
'#Uses "*bsShowMessage"

Public Sub BOTAOPROCESSAR_OnClick()
	If ((CurrentQuery.State = 2) Or _
		(CurrentQuery.State = 3)) Then
		bsShowMessage("O registro não pode estar em edição ou inclusão!", "I")

		Exit Sub
	End If
	If CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
		bsShowMessage("Rotina já processada", "I")

		Exit Sub
	End If
	If VisibleMode Then
		Dim vvInterface As Object
		Set vvInterface = CreateBennerObject("BSINTERFACE0051.Rotinas")

		vvInterface.ImportarAbramge(CurrentSystem, _
									CurrentQuery.FieldByName("HANDLE").AsInteger, _
									CurrentQuery.FieldByName("FILIAL").AsInteger)
	Else
		Dim vsMensagemErro As String
		Dim viRetorno As Long
		Dim obj As Object
		Dim vcContainer As CSDContainer
		Set vcContainer = NewContainer

		vcContainer.AddFields("HANDLE:INTEGER")
		vcContainer.AddFields("FILIAL:INTEGER")
		vcContainer.Insert

		vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		vcContainer.Field("FILIAL").AsInteger = CurrentQuery.FieldByName("FILIAL").AsInteger

		Set obj = CreateBennerObject("BSSERVEREXEC.ProcessosServidor")

		viRetorno = obj.ExecucaoImediata(CurrentSystem, _
										 "SAMPEGDIGIT", _
										 "ImportarAbramge", _
										 "Importação de arquivo ABRAMGE", _
										 0, _
										 "SAM_ROTINAABRAMGE", _
										 "SITUACAO", _
										 "", _
										 "", _
										 "P", _
										 True, _
										 vsMensagemErro, _
										 vcContainer)

		If viRetorno = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
		Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		End If
	End If

	Set obj = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
		bsShowMessage("Rotina já processada", "I")

		CanContinue = False
		Exit Sub
	End If
	CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOPROCESSAR" Then
		BOTAOPROCESSAR_OnClick
	End If
End Sub
