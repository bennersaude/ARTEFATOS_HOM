'HASH: C4A9F30329734F2D532F9160196556FD
'SAM_LIVROCONFIG_FILTROPREST
'#Uses "*bsShowMessage"

Public Sub BOTAOGERARPRESTADORES_OnClick()
	Dim Interface As Object
	Dim Tabela, CampoStr As String
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "I")
		CanContinue = False
		Exit Sub
	End If

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "I")
		CanContinue = False
		Exit Sub
	End If

	Tabela = "SAM_LIVROCONFIG_FILTROPREST"
	CampoStr = "LIVROCONFIGURACAO"

	Set Interface = CreateBennerObject("BSPRE001.Rotinas")

	Interface.GerarPrestadores(CurrentSystem, Tabela, CampoStr, CurrentQuery.FieldByName("LIVROCONFIGURACAO").AsInteger)

	RefreshNodesWithTable("SAM_LIVROCONFIG_FILTROPREST")

	Set Interface = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim Handlexx As Long
	Dim vCondicao As String

	ShowPopup = False

	Set Interface = CreateBennerObject("Procura.Procurar")

	Handlexx = -1
	vCondicao = ""
	Handlexx = Interface.Exec(CurrentSystem, "SAM_PRESTADOR", "PRESTADOR|NOME", 1, "Prestador|Nome", vCondicao, "Tabela de prestadores", True, "")

	If Handlexx > 0 Then
		CurrentQuery.FieldByName("PRESTADOR").Value = Handlexx
	End If

	Set Interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_LIVROCONFIG_FILTROPREST WHERE PRESTADOR = :PRESTADOR AND LIVROCONFIGURACAO = :LIVROCONFIGURACAO AND HANDLE <> :HANDLE")

	SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
	SQL.ParamByName("LIVROCONFIGURACAO").Value = CurrentQuery.FieldByName("LIVROCONFIGURACAO").Value
	SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
	SQL.Active = True

	If Not SQL.EOF Then
		bsShowMessage("Registro duplicado para esta configuração !", "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQL = Nothing
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
		Case "BOTAOGERARPRESTADORES"
			BOTAOGERARPRESTADORES_OnClick
	End Select
End Sub
