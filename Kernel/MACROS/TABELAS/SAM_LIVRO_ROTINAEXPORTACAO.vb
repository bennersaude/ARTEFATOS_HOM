'HASH: 2CE57137C12C4D5963432B4D9847F203
'SAM_LIVRO_ROTINAEXPORTACAO
'#Uses "*bsShowMessage"

Public Sub BOTAOLOCAL_OnClick()
	Dim Interface As Object
	Dim vPath As String
	Set Interface = CreateBennerObject("BSPRE001.Rotinas")

	vPath = Interface.SelecionarDiretorio(CurrentSystem)

	If vPath <> "" Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("LOCALARQUIVO").AsString = vPath
	End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
	Dim Obj As Object
	Dim vOK As Integer
	Dim SQL As Object
	Set SQL = NewQuery

	If CurrentQuery.FieldByName("TIPOEXPORTACAO").AsString = "E" Then
		SQL.Add("SELECT * FROM SAM_LIVROENCARTE WHERE LIVRO = :LIVRO")
		SQL.ParamByName("LIVRO").Value = CurrentQuery.FieldByName("LIVRO").Value
		SQL.Active = True

		If SQL.EOF Then
			bsShowMessage("Este livro não possue encartes !", "I")
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	'SMS 90283 - Ricardo Rocha - Adequação WEB
	If VisibleMode Then
    	Set Obj = CreateBennerObject("BSInterface0007.Rotinas")
    	Obj.ExportarDados(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  	Else
    	Dim vsMensagemErro As String
    	Dim viRetorno As Long

    	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSPRE001", _
                                     "Rotinas_ExportarDados", _
                                     "Exportacao de Livro de Credenciados - Rotina: " + _
                                     CStr(CurrentQuery.FieldByName("ROTINA").AsInteger) + _
                                     " Livro: " + CStr(CurrentQuery.FieldByName("LIVRO").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_LIVRO_ROTINAEXPORTACAO", _
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

	RefreshNodesWithTable("SAM_LIVRO_ROTINAEXPORTACAO")

End Sub

Public Sub LIVRO_OnPopup(ShowPopup As Boolean)
	CurrentQuery.FieldByName("LIVROENCARTE").Clear
End Sub

Public Sub TABLE_AfterScroll()
	LIVROENCARTE.Visible = True

	vCondicao = ""

	If VisibleMode Then
		vCondicao = vCondicao + "SAM_LIVROENCARTE.HANDLE "
		vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_LIVROENCARTE WHERE LIVRO = @LIVRO)"

		LIVROENCARTE.LocalWhere = vCondicao
	Else
		VCONDICAO = VCONDICAO + "A.HANDLE "
		vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_LIVROENCARTE WHERE LIVRO = @CAMPO(LIVRO))"

		LIVROENCARTE.WebLocalWhere = vCondicao
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("TIPOEXPORTACAO").AsString = "E" And CurrentQuery.FieldByName("LIVROENCARTE").IsNull Then
		bsShowMessage("Campo Encarte obrigatório !", "E")
		LIVROENCARTE.Visible = True
		CanContinue = False
		Exit Sub
	Else
		If CurrentQuery.FieldByName("TIPOEXPORTACAO").AsString = "C" And Not CurrentQuery.FieldByName("LIVROENCARTE").IsNull Then
			CurrentQuery.FieldByName("LIVROENCARTE").Value = Null
		End If
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOLOCAL"
			BOTAOLOCAL_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
