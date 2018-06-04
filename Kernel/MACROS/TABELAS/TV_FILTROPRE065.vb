'HASH: EEC72D205B0D1665753F134A7DA3BA05
 
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTROPRE065") <> "" Then
		XMLToDataset(UserVar("TV_FILTROPRE065"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	Dim dataInicial As Date

	UserVar("TV_FILTROPRE065") = DatasetToXML(CurrentQuery.TQuery,"")

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		dataInicial = (CurrentQuery.FieldByName("DATAFINAL").AsDateTime - CurrentQuery.FieldByName("DIASPESQUISA").AsInteger)

		Dim vFiltro As String

		vFiltro = "P.DATADESCREDENCIAMENTO BETWEEN '" &  dataInicial & "' AND '" & CurrentQuery.FieldByName("DATAFINAL").AsDateTime & "'" & _
				  IIf(CurrentQuery.FieldByName("FILIAIS").AsString <> "", " AND " & FilterFieldResultSQL("F.HANDLE",CurrentQuery.FieldByName("FILIAIS").AsString), "") & _
				  IIf(CurrentQuery.FieldByName("TIPOPRESTADOR").AsString <> "", " AND " & FilterFieldResultSQL("TP.HANDLE",CurrentQuery.FieldByName("TIPOPRESTADOR").AsString), "")


		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)


		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - PRE065 - Descredenciados por intervalo de dias"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "PRE065"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Prestadores"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
