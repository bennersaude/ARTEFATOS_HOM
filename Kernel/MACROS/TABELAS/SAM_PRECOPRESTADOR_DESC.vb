'HASH: 498D56C58FF54E1746FCB07C0E85BC98
'Macro: SAM_PRECOPRESTADOR_DESC
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String
    Dim vNumeroColuna As Integer

    ShowPopup = False

    Set interface =CreateBennerObject("Procura.Procurar")

    vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
      vColunas ="ESTRUTURA|ESTRUTURANUMERICA|DESCRICAOABREVIADA|NIVELAUTORIZACAO"
      vCampos ="Estrutura|Estrutura Numérica|Descrição|Nível de autorização"

    If IsNumeric(EVENTOINICIAL.Text) Then
        vNumeroColuna = 2

    Else
        vNumeroColuna = 3

    End If

    vHandle =interface.Exec(CurrentSystem,"SAM_TGE",vColunas, vNumeroColuna, vCampos, vCriterio, "Evento", False, EVENTOINICIAL.Text)

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
    End If

    Set interface =Nothing
End Sub

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String
    Dim vNumeroColuna As Integer

    ShowPopup = False

    Set interface =CreateBennerObject("Procura.Procurar")

    vCriterio = "SAM_TGE.ULTIMONIVEL = 'S'"
      vColunas ="ESTRUTURA|ESTRUTURANUMERICA|DESCRICAOABREVIADA|NIVELAUTORIZACAO"
      vCampos ="Estrutura|Estrutura Numérica|Descrição|Nível de autorização"

    If IsNumeric(EVENTOFINAL.Text) Then
        vNumeroColuna = 2

    Else
        vNumeroColuna = 3

    End If

    vHandle =interface.Exec(CurrentSystem,"SAM_TGE",vColunas, vNumeroColuna, vCampos, vCriterio, "Evento", False, EVENTOFINAL.Text)

    If vHandle <>0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
    End If

    Set interface =Nothing
End Sub

Public Sub TABLE_AfterEdit()
	UpdateLastUpdate("SAM_CONVENIO")

	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim interface As Object
	Dim Linha As String
	Dim Condicao As String
	Dim EstruturaI As String
	Dim EstruturaF As String
	' Atribuir ESTRUTURAINICIAL E FINAL
	Dim SQLTGE, SQLMASC As Object
	Dim Estrutura As String
	' Atribuir ESTRUTURAINICIAL
	Set SQLTGE = NewQuery
	Dim EspecificoDll As Object

	SQLTGE.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTO")

	SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
	SQLTGE.Active = True

	CurrentQuery.FieldByName("ESTRUTURAINICIAL").Value = SQLTGE.FieldByName("ESTRUTURA").Value

	' Atribuir ESTRUTURAFINAL
	SQLTGE.Active = False
	SQLTGE.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
	SQLTGE.Active = True

	Estrutura = SQLTGE.FieldByName("ESTRUTURA").Value

	SQLTGE.Active = False

	Set SQLTGE = Nothing
	' Completar ESTRUTURAFinal com 99999
	Set SQLMASC = NewQuery

	SQLMASC.Add("SELECT M.MASCARA MASCARA FROM Z_TABELAS T, Z_MASCARAS M")
	SQLMASC.Add("WHERE T.NOME = 'SAM_TGE' AND M.TABELA = T.HANDLE")

	SQLMASC.Active = True

	Estrutura = Estrutura + Mid(SQLMASC.FieldByName("MASCARA").AsString, Len(Estrutura) + 1)

	CurrentQuery.FieldByName("ESTRUTURAFINAL").Value = Estrutura

	SQLMASC.Active = False

	Set SQLMASC = Nothing

	If CanContinue = True Then
		' Checar Vigencia
		EstruturaI = CurrentQuery.FieldByName("ESTRUTURAINICIAL").AsString
		EstruturaF = CurrentQuery.FieldByName("ESTRUTURAFINAL").AsString
		Condicao = " PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString
		Condicao = Condicao + " AND CLASSEASSOCIADO = '" + CurrentQuery.FieldByName("CLASSEASSOCIADO").AsString + "'"

		If CurrentQuery.FieldByName("CONVENIO").IsNull Then
			Condicao = Condicao + " AND CONVENIO IS NULL"
		Else
			Condicao = Condicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
		End If

		Set EspecificoDll = CreateBennerObject("ESPECIFICO.UESPECIFICO")
	    Condicao = Condicao + EspecificoDll.CAM_PRO_VerificarVigenciaDescontoPrecoPrestador(CurrentSystem, CurrentQuery.TQuery)

		Set interface = CreateBennerObject("SAMGERAL.Vigencia")

		Linha = interface.EventoFx(CurrentSystem, "SAM_PRECOPRESTADOR_DESC", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, EstruturaI, EstruturaF, Condicao)

		If Linha = "" Then
			CanContinue = True
		Else
			CanContinue = False
			bsShowMessage(Linha, "E")
			Exit Sub
		End If

		Set interface = Nothing
		Set EspecificoDll = Nothing
	End If

	If CanContinue = True Then
		CanContinue = CheckEventosFx
	End If
End Sub

Public Function CheckEventosFx As Boolean
	CheckEventosFx = True

	If Not CurrentQuery.FieldByName("EVENTOINICIAL").IsNull Then
		If CurrentQuery.FieldByName("EVENTOFINAL").IsNull Then
			CurrentQuery.FieldByName("EVENTOFINAL").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
		Else
			If CurrentQuery.FieldByName("EVENTOINICIAL").Value <>CurrentQuery.FieldByName("EVENTOFINAL").Value Then
				Dim SQLI, SQLF As Object
				Set SQLI = NewQuery

				SQLI.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOI")

				SQLI.ParamByName("HEVENTOI").Value = CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger
				SQLI.Active = True

				Set SQLF = NewQuery

				SQLF.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE = :HEVENTOF")

				SQLF.ParamByName("HEVENTOF").Value = CurrentQuery.FieldByName("EVENTOFINAL").AsInteger
				SQLF.Active = True

				If SQLF.FieldByName("ESTRUTURA").Value <SQLI.FieldByName("ESTRUTURA").Value Then
					bsShowMessage("Evento final não pode ser menor que o evento inicial!", "E")
					EVENTOFINAL.SetFocus
					CheckEventosFx = False
				End If

				Set SQLI = Nothing
				Set SQLF = Nothing
			End If
		End If
	End If
End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
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

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

	SQL.Active = True

	If SQL.FieldByName("TOTAL").AsInteger = 1 Then
		SQL.Active = False

		SQL.Clear

		SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

		SQL.Active = True

		CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
	End If

	Set SQL = Nothing

	UpdateLastUpdate("SAM_CONVENIO")

	Dim vCondicao As String

	If VisibleMode Then
		vCondicao = "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = "A.HANDLE "
	End If

	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = CONVENIOMESTRE)"

	If VisibleMode Then
		CONVENIO.LocalWhere = vCondicao
	Else
		CONVENIO.WebLocalWhere = vCondicao
	End If
End Sub
