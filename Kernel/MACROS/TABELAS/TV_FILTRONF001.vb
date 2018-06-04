'HASH: 065472EE605FB27DC9A187BBE94CE4DF
'#uses "*Biblioteca"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FILTRONF001") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FILTRONF001"),CurrentQuery.TQuery)
	End If
End Sub


Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FILTRONF001") = DatasetToXML(CurrentQuery.TQuery,"")

	Call CriarFiltro("NF001", CurrentUser, "CHECKGERAL= '" & CurrentQuery.FieldByName("IMPRIMIRRESUMO").AsString & "'" , _
	                                         "CHECKGERAL", _
	                                         "'" & CurrentQuery.FieldByName("IMPRIMIRRESUMO").AsString & "'")


	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String



		vFiltro = "CP.COMPETENCIA = " & SQLDate(CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)

		If CurrentQuery.FieldByName("PRESTADOR").AsString <> "" Then
			vFiltro = vFiltro & "AND P.HANDLE = " & CurrentQuery.FieldByName("PRESTADOR").AsString
		End If


		If CurrentQuery.FieldByName("FILIAL").AsString <> "" Then
			vFiltro = vFiltro & "AND FL.HANDLE = " & CurrentQuery.FieldByName("FILIAL").AsString
		End If

		If CurrentQuery.FieldByName("MUNICIPIO").AsString <> "" Then
			vFiltro = vFiltro & "AND M.HANDLE = " & CurrentQuery.FieldByName("MUNICIPIO").AsString
		End If

		If CurrentQuery.FieldByName("ESTADO").AsString <> "" Then
			vFiltro = vFiltro & "AND E.HANDLE = " & CurrentQuery.FieldByName("ESTADO").AsString
		End If

		vFiltro = vFiltro & " AND ((P.FISICAJURIDICA = '1' AND '"& CurrentQuery.FieldByName("TIPOPESSOA").AsString & "' = 'F') OR (P.FISICAJURIDICA = '2' AND  '" & CurrentQuery.FieldByName("TIPOPESSOA").AsString &"' = 'J') OR ('A' = '" & CurrentQuery.FieldByName("TIPOPESSOA").AsString & "'))"


		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)


		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - NF001 - Relação de Notas Fiscais por Prestador"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "NF001"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Prestadores"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
