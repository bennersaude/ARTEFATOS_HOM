'HASH: 1F4DF6C05BE315C756373F1A718ED91A
 
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTRO_PRE068") <> "" Then
		XMLToDataset(UserVar("TV_FILTRO_PRE068"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	UserVar("TV_FILTRO_PRE068") = DatasetToXML(CurrentQuery.TQuery,"")

	Dim vDataInicial As String
	Dim vDataFinal As String

	If CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
		vDataInicial = "NULL"
	Else
		vDataInicial = SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
	End If


	If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		vDataFinal   = "NULL"
	Else
		vDataFinal   = SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)
	End If

	Call CriarFiltro("PRE068", CurrentUser, "DATAINICIAL=" & vDataInicial & ", DATAFINAL= " & vDataFinal, _
	                                         "DATAINICIAL, DATAFINAL", _
	                                         vDataInicial & ", " & vDataFinal)

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String
        vFiltro = "1 = 1"

		If CurrentQuery.FieldByName("ESTADO").AsInteger > 0 Then
			vFiltro = vFiltro & " AND F.HANDLE = " & CurrentQuery.FieldByName("ESTADO").AsInteger
		End If

		If CurrentQuery.FieldByName("MUNICIPIO").AsInteger > 0 Then
			vFiltro = vFiltro & " AND E.HANDLE = " & CurrentQuery.FieldByName("MUNICIPIO").AsInteger
		End If

       	If CurrentQuery.FieldByName("PRESTADOR").AsInteger > 0 Then
			vFiltro = vFiltro & " AND B.HANDLE = " & CurrentQuery.FieldByName("PRESTADOR").AsInteger
		End If

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - PRE068 - Histórico de documentos entregues dos prestadores"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "PRE068"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Prestadores"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
