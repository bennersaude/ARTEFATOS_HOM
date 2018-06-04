'HASH: AC76BDEB417ABBC875265567614C4326
'SAM_LIVROCONFIG_FILTROMUNIC
'#Uses "*bsShowMessage"

Public Sub MUNICIPIO_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim Handlexx As Long
	Dim vCondicao As String
	Dim SQL As Object

	ShowPopup = False

	Set Interface = CreateBennerObject("Procura.Procurar")

	Handlexx = -1

	Set SQL = NewQuery

	SQL.Add("SELECT ESTADO FROM SAM_LIVROCONFIG_FILTROESTADO WHERE LIVROCONFIGURACAO = :LIVROCONFIGURACAO")

	SQL.ParamByName("LIVROCONFIGURACAO").Value = CurrentQuery.FieldByName("LIVROCONFIGURACAO").Value
	SQL.Active = True

	vCondicao = ""

	If Not SQL.EOF Then
		vCondicao = vCondicao + "(ESTADOS.HANDLE = " + SQL.FieldByName("ESTADO").AsString

		SQL.Next

		While Not SQL.EOF
			vCondicao = vCondicao + "  OR ESTADOS.HANDLE = " + SQL.FieldByName("ESTADO").AsString
			SQL.Next
		Wend

		vCondicao = vCondicao + ")"
	End If

	Handlexx = Interface.Exec(CurrentSystem, "MUNICIPIOS|ESTADOS[ESTADOS.HANDLE = MUNICIPIOS.ESTADO]", "MUNICIPIOS.NOME|ESTADOS.SIGLA", 1, "Municípios|Estado", vCondicao, "Tabela de municipios", True, "")

	If Handlexx > 0 Then
		CurrentQuery.FieldByName("MUNICIPIO").Value = Handlexx
	End If

	Set SQL = Nothing
	Set Interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT * FROM SAM_LIVROCONFIG_FILTROMUNIC WHERE MUNICIPIO = :MUNICIPIO AND LIVROCONFIGURACAO = :LIVROCONFIGURACAO AND HANDLE <> :HANDLE")

	SQL.ParamByName("MUNICIPIO").Value = CurrentQuery.FieldByName("MUNICIPIO").Value
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
