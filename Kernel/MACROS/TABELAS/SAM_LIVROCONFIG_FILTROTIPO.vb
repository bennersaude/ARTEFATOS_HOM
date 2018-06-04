'HASH: B1AC27BF06EAEDC3F16FD95CA1DD74EF
'SAM_LIVROCONFIG_FILTROTIPO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim regDuplicado As BPesquisa
	Set regDuplicado = NewQuery
	regDuplicado.Clear
	regDuplicado.Add("SELECT COUNT(1) QTD                   ")
	regDuplicado.Add("  FROM SAM_LIVROCONFIG_FILTROTIPO     ")
	regDuplicado.Add(" WHERE HANDLE <> :HANDLE              ")
	regDuplicado.Add("   AND LIVROCONFIGURACAO = :LIVROCONFIGURACAO ")
	regDuplicado.Add("   AND TIPOPRESTADOR = :TIPOPRESTADOR ")
    regDuplicado.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    regDuplicado.ParamByName("TIPOPRESTADOR").Value = CurrentQuery.FieldByName("TIPOPRESTADOR").AsInteger
    regDuplicado.ParamByName("LIVROCONFIGURACAO").Value = CurrentQuery.FieldByName("LIVROCONFIGURACAO").AsInteger
	regDuplicado.Active = True

	If regDuplicado.FieldByName("QTD").AsInteger > 0 Then
		Set regDuplicado = Nothing
		bsShowMessage("Tipo de prestador já cadastrado para essa configuração", "E")
		CanContinue = False
	End If

	Set regDuplicado = Nothing
End Sub

Public Sub TIPOPRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim Handlexx As Long
	Dim vCondicao As String

	ShowPopup = False

	Set Interface = CreateBennerObject("Procura.Procurar")

	Handlexx = -1
	vCondicao = ""
	Handlexx = Interface.Exec(CurrentSystem, "SAM_TIPOPRESTADOR", "DESCRICAO", 1, "Descrição", vCondicao, "Tabela de tipo de prestadores", True, "")

	If Handlexx > 0 Then
		CurrentQuery.FieldByName("TIPOPRESTADOR").Value = Handlexx
	End If

	Set Interface = Nothing
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
