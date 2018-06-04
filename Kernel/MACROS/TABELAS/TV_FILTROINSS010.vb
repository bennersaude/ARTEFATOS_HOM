'HASH: 6470F5E2283A5947FD41D2D8D9ECD65F
 
'#uses "*Biblioteca"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FILTROINSS010") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FILTROINSS010"),CurrentQuery.TQuery)
	End If
End Sub


Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FILTROINSS010") = DatasetToXML(CurrentQuery.TQuery,"")

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String

		If InStr(SQLServer,"ORACLE") > 0 Then
			vFiltro = "CP.COMPETENCIA = TO_DATE('" & CurrentQuery.FieldByName("COMPETENCIA").AsDateTime & "','DD/MM/YYYY')"
		Else
			vFiltro = "CP.COMPETENCIA = '" & FormatDateTime2("YYYY-MM-DD",CurrentQuery.FieldByName("COMPETENCIA").AsDateTime) & "'"
		End If

		If CurrentQuery.FieldByName("PRESTADOR").AsString <> "" Then
			vFiltro = vFiltro & "AND P.HANDLE = " & CurrentQuery.FieldByName("PRESTADOR").AsString
		End If

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "DECLARAÇÃO DE PAGAMENTO A CONTRIBUINTE INDIVIDUAL E COOPERADO"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "INSS010"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Controle Financeiro"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
