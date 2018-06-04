'HASH: 0456AE9D9C7DD5E3253D8A4D0CB37E69
'#uses "*Biblioteca"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTRO0192") <> "" Then
		XMLToDataset(UserVar("TV_FILTRO0192"),CurrentQuery.TQuery)
	End If
End Sub


Public Sub TABLE_AfterPost()
	UserVar("TV_FILTRO0192") = DatasetToXML(CurrentQuery.TQuery,"")

	UserVar("DATAVENCIMENTOTVFILTRO0192") = IIf(Not CurrentQuery.FieldByName("DATAVENCIMENTO").IsNull, CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime, "")

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String

		vFiltro = IIf(Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull, " B.HANDLE IN (" & ConvertPipeToVirgulaCampoFiltro(CurrentQuery.FieldByName("BENEFICIARIO").AsString) & ")", "")

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "SFN004 - Guia de Recolhimento da União (GRU)"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "SFN004"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Controle Financeiro"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute

		Set sx = Nothing
		Set container = Nothing

	End If
End Sub
