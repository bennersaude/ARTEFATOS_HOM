'HASH: 0EF63F071BCED65D686AEF6E72150605
'Macro: SAM_PRESTADOR_LIVRO
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterEdit()
	UpdateLastUpdate("SAM_ESPECIALIDADE")

	Dim vSelect As String

	vSelect = "(SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE PRESTADOR = "

	If VisibleMode Then
		ESPECIALIDADE.LocalWhere ="HANDLE IN " + vSelect + "@PRESTADOR)"
	Else
		ESPECIALIDADE.WebLocalWhere ="A.HANDLE IN " + vSelect + "@CAMPO(PRESTADOR))"
	End If
End Sub

Public Sub TABLE_AfterInsert()
	UpdateLastUpdate("SAM_ESPECIALIDADE")

	Dim vSelect As String

	vSelect = "(SELECT ESPECIALIDADE FROM SAM_PRESTADOR_ESPECIALIDADE WHERE PRESTADOR = "

	If VisibleMode Then
		ESPECIALIDADE.LocalWhere ="HANDLE IN " + vSelect + "@PRESTADOR)"
	Else
		ESPECIALIDADE.WebLocalWhere ="A.HANDLE IN " + vSelect + "@CAMPO(PRESTADOR))"
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If Not LivroJaExiste() Then
		Dim SQL As Object
		Set SQL = NewQuery

		SQL.Add("SELECT COUNT(*) CONTADOR FROM SAM_DIMENSIONAMENTO ")
		SQL.Add("WHERE AREALIVRO = :AREA AND ESPECIALIDADE = :ESPECIALIDADE ")

		SQL.ParamByName("ESPECIALIDADE").Value  = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
		SQL.ParamByName("AREA").Value = CurrentQuery.FieldByName("AREA").AsInteger
		SQL.Active = True

		If SQL.FieldByName("CONTADOR").AsInteger = 0 Then
			bsShowMessage("Especialidade não cadastrada para esta área!", "E")
			CanContinue = False
			Set SQL = Nothing
			Exit Sub
		End If

		Dim vData As String
		Dim qESP As Object
		Set qESP = NewQuery

		vData = SQLDate( ServerDate)

		qESP.Add("SELECT * FROM SAM_PRESTADOR_ESPECIALIDADE WHERE ESPECIALIDADE = :ESPECIALIDADE ")
		qESP.Add("                                            AND PRESTADOR     = :PRESTADOR     ")
		qESP.Add("                                            AND DATAINICIAL <= "+vData+"       ")
		qESP.Add("                                            AND (DATAFINAL IS NULL OR DATAFINAL >= "+vData+")")

		qESP.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").Value
		qESP.ParamByName("PRESTADOR").Value     = CurrentQuery.FieldByName("PRESTADOR").Value
		qESP.Active=True

		If CurrentQuery.FieldByName("PUBLICARINTERNET").AsString = "S" And qESP.FieldByName("PUBLICARINTERNET").AsString = "N" Then
			bsShowMessage("Campo 'Publicar no internet' não pode ser marcado !" + Chr(10) + _
				"Motivo: A especialidade está cadastrada no prestador com este campo desmarcado.", "E")
			CurrentQuery.FieldByName("PUBLICARINTERNET").Value = "N"
		End If

		If CurrentQuery.FieldByName("VISUALIZARCENTRAL").AsString = "S" And qESP.FieldByName("VISUALIZARCENTRAL").AsString = "N" Then
			bsShowMessage("Campo 'Visualizar na central de atendimento' não pode ser marcado !" + Chr(10) + _
				"Motivo: A especialidade está cadastrada no prestador com este campo desmarcado.", "E")
			CurrentQuery.FieldByName("VISUALIZARCENTRAL").Value = "N"
		End If
	Else
		bsShowMessage("Registro de credenciado já existe.","E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E","P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A","P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I","P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
Public Function LivroJaExiste()As Boolean
Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) CONTADOR FROM SAM_PRESTADOR_LIVRO  ")
	SQL.Add("WHERE AREA = :AREA AND ESPECIALIDADE = :ESPECIALIDADE AND ENDERECO  = :ENDERECO AND HANDLE <> :HANDLE")
	SQL.ParamByName("ESPECIALIDADE").Value  = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
	SQL.ParamByName("AREA").Value = CurrentQuery.FieldByName("AREA").AsInteger
	SQL.ParamByName("ENDERECO").Value  = CurrentQuery.FieldByName("ENDERECO").AsInteger
	SQL.ParamByName("HANDLE").Value  = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.Active = True

    LivroJaExiste = SQL.FieldByName("CONTADOR").AsInteger > 0
End Function
