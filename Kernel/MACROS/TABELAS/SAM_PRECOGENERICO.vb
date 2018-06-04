'HASH: A8CCAB30725C24F28B143DBE6281682A
'Macro: SAM_PRECOGENERICO
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*bsShowMessage

Option Explicit

Public Sub BOTAOIMPORTACAO_OnClick()
	Dim Obj As Object
	Dim mensagem As String

	If CurrentQuery.State <>1 Then
		bsShowMessage("A tabela não pode estar em edição", "I")
		Exit Sub
	End If


    If VisibleMode Then
        Set Obj = CreateBennerObject("BSINTERFACE0008.Rotinas")

    	Obj.GERAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	    Set Obj = Nothing

    Else
		Dim vsMensagemErro As String
	  	Dim viRet As Long

		Dim vcContainer As CSDContainer
	   	Set vcContainer = NewContainer
	   	vcContainer.AddFields("HANDLE:INTEGER")

		vcContainer.Insert
	 	vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger


	  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		viRet = Obj.ExecucaoImediata(CurrentSystem, _
		                           	 "SAMImportaTabPrc", _
			                         "ImportarTabPrc", _
			                         "Importação de tabela de preço", _
			                         CurrentQuery.FieldByName("HANDLE").AsInteger, _
			                         "SAM_PRECOGENERICO", _
			                         "SITUACAO", _
			                         "", _
			                         "", _
			                         "P", _
			                         True, _
			                         vsMensagemErro, _
			                         vcContainer)

		If viRet = 0 Then
		 	bsShowMessage("Processo enviado para execução no servidor!", "I")
		Else
	     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	   	End If

		Set Obj = Nothing
	End If
End Sub

Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vColunas As String
	Dim vCriterio As String
	Dim vCampos As String
	Set Interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vColunas = "DESCRICAO"
	vCriterio = ""
	vCampos = "Descrição da Tabela"

	CurrentQuery.FieldByName("TABELAUS").Value = Interface.Exec(CurrentSystem, "SAM_TABUS", vColunas, 1, vCampos, vCriterio, "Tabela de US", True, TABELAUS.Text)

	Set Interface = Nothing
End Sub

Public Sub TABELAFILME_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraTabelaFilme(TABELAFILME.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("TABELAFILME").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterPost()
	RefreshNodesWithTable("SAM_PRECOGENERICO")
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOIMPORTACAO"
			BOTAOIMPORTACAO_OnClick
	End Select
End Sub
