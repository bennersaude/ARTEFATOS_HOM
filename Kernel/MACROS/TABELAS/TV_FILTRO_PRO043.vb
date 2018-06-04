'HASH: 5DBCE5B5329E346E126592685128F1F2
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTRO_PRO043") <> "" Then
		XMLToDataset(UserVar("TV_FILTRO_PRO043"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	UserVar("TV_FILTRO_PRO043") = DatasetToXML(CurrentQuery.TQuery,"")

	Dim vDataInicial As String
	Dim vDataFinal As String

	vDataInicial = SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
	vDataFinal   = SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

	Call CriarFiltro("PRO043", CurrentUser, "DATAINICIAL=" & vDataInicial & ", DATAFINAL= " & vDataFinal, _
	                                         "DATAINICIAL, DATAFINAL", _
	                                         vDataInicial & ", " & vDataFinal)

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String
        vFiltro = "1 = 1"

		If CurrentQuery.FieldByName("FILIAL").AsInteger > 0 Then
			vFiltro = vFiltro & " AND A.FILIAL = " & CurrentQuery.FieldByName("FILIAL").AsInteger
		End If

		If CurrentQuery.FieldByName("RECEBEDOR").AsInteger > 0 Then
			vFiltro = vFiltro & " AND A.RECEBEDOR = " & CurrentQuery.FieldByName("RECEBEDOR").AsInteger
		End If

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - PRO043 - Extrato de eventos pagos sem autorização vinculada"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "PRO043"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Processamento de Contas"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
