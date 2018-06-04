'HASH: C0870176C2E653343CCEAED59359C939
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FORM_BEN085") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FORM_BEN085"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	Dim DataAno As String

	DataAno =  CStr(Year(CurrentQuery.FieldByName("ANO").AsDateTime))

	Call CriarFiltro("BEN085", CurrentUser, " DATAINICIAL = '01/01/" & DataAno & "' ,"  & " DATAFINAL = '31/12/" & DataAno & "'" , _
	                 "DATAINICIAL, DATAFINAL", "'01/01/" & DataAno & "', '31/12/" & DataAno & "'")


	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String

		UserVar("FILTRO_TV_FORM_BEN085") = DatasetToXML(CurrentQuery.TQuery,"")

		vFiltro = "1=1" & IIf(CurrentQuery.FieldByName("FILIAL").AsString<>""," AND " &  FilterFieldResultSQL("F.HANDLE",CurrentQuery.FieldByName("FILIAL").AsString), "") & _
				          IIf(CurrentQuery.FieldByName("VINCULOFUNCIONAL").AsString<>""," AND " &  FilterFieldResultSQL("I.HANDLE",CurrentQuery.FieldByName("VINCULOFUNCIONAL").AsString),"") & _
				          IIf(CurrentQuery.FieldByName("BENEFICIARIO").AsString<>""," AND " &  FilterFieldResultSQL("B.HANDLE",CurrentQuery.FieldByName("BENEFICIARIO").AsString),"")



		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)


		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - BEN085 - Beneficiários com utilização superior ao teto anual"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "BEN085"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Beneficiários"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If
End Sub
