'HASH: 07A12DDF326DC2D0D91FA1C645391E06

'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTRO_PRO040") <> "" Then
		XMLToDataset(UserVar("TV_FILTRO_PRO040"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	UserVar("TV_FILTRO_PRO040") = DatasetToXML(CurrentQuery.TQuery,"")

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String

		If InStr(SQLServer,"ORACLE") > 0 Then
			vFiltro = "AP.DATAPAGAMENTO = TO_DATE('" & CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime & "','DD/MM/YYYY')"
		Else
			vFiltro = "AP.DATAPAGAMENTO = '" & FormatDateTime2("YYYY-MM-DD", CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime) & "'"
		End If

		If CurrentQuery.FieldByName("RECEBEDOR").AsInteger > 0 Then
			vFiltro = vFiltro & " AND P.HANDLE = " & CurrentQuery.FieldByName("RECEBEDOR").AsInteger
		End If

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - PRO040 - RPA por Data de Pagamento"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "PRO040"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Processamento de Contas"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
 
