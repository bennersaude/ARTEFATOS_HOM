'HASH: 2901A5B7EA2DBF6E84EE0CB69A853DA0
 
'#uses "*Biblioteca"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FILTROIMP002") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FILTROIMP002"),CurrentQuery.TQuery)
	End If
End Sub


Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FILTROIMP002") = DatasetToXML(CurrentQuery.TQuery,"")

	Call CriarFiltro("IMP002", CurrentUser, "DATAINICIAL ='" & CurrentQuery.FieldByName("DATAINICIAL").AsDateTime & "',"  & "DATAFINAL='" & CurrentQuery.FieldByName("DATAFINAL").AsString & "'", "DATAINICIAL, DATAFINAL", "'" & CurrentQuery.FieldByName("DATAINICIAL").AsDateTime & "','" & CurrentQuery.FieldByName("DATAFINAL").AsString & "'" )

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String

		If InStr(SQLServer,"ORACLE") > 0 Then
			vFiltro = "AP.DATAPAGAMENTO BETWEEN TO_DATE('" & CurrentQuery.FieldByName("DATAINICIAL").AsDateTime & "','DD/MM/YYYY') AND TO_DATE('" & CurrentQuery.FieldByName("DATAFINAL").AsDateTime & "','DD/MM/YYYY')"
		Else
			vFiltro = "AP.DATAPAGAMENTO BETWEEN '" & FormatDateTime2("YYYY-MM-DD",CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) & "' AND '" & FormatDateTime2("YYYY-MM-DD",CurrentQuery.FieldByName("DATAFINAL").AsDateTime) & "'"
		End If


		If CurrentQuery.FieldByName("FILIAL").AsString <> "" Then
			vFiltro = vFiltro & "AND F.HANDLE = " & CurrentQuery.FieldByName("FILIAL").AsString
		End If


		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)


		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - IMP002 - Resumo de Impostos por Agrupador de Pagamento"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "IMP002"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Controle Financeiro"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
