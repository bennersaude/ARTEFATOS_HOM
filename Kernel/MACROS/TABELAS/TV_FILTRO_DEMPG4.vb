'HASH: 5B963760A7B8B0F218F32266CAC65B76
 

'#uses "*Biblioteca"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FILTRODEMPG4") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FILTRODEMPG4"),CurrentQuery.TQuery)
	End If
End Sub


Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FILTRODEMPG4") = DatasetToXML(CurrentQuery.TQuery,"")

	Dim vDataInicial As String
	Dim vDataFinal As String

	vDataInicial = SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
	vDataFinal   = SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)

	Call CriarFiltro("DEM-PG4", CurrentUser, "FILIAL= " & CurrentQuery.FieldByName("FILIAL").AsString & ", DATAINICIAL=" & vDataInicial & ", DATAFINAL= " & vDataFinal & ", PRESTADOR= " & IIf(CurrentQuery.FieldByName("PRESTADOR").AsInteger=0,"NULL",CurrentQuery.FieldByName("PRESTADOR").AsString), _
	                                         "FILIAL, DATAINICIAL, DATAFINAL, PRESTADOR", _
	                                         CurrentQuery.FieldByName("FILIAL").AsString & ", " & vDataInicial & ", " & vDataFinal & ", " & IIf(CurrentQuery.FieldByName("PRESTADOR").AsInteger=0,"NULL",CurrentQuery.FieldByName("PRESTADOR").AsString))

	If WebMode Then

		Dim container As CSDContainer
		Set container = NewContainer

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - DEM-PG4-Demonstrativo de Pagamento Prestador"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "DEM-PG4"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Controle Financeiro"))
		sx.SessionVar("DefaultWhere") = ""
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If

End Sub
