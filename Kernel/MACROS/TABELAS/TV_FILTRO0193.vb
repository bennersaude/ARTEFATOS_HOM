'HASH: 0B1F3087D47AA025BA70C2E945C6F457
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FORM_COM001") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FORM_COM001"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FORM_COM001") = DatasetToXML(CurrentQuery.TQuery,"")

	UserVar("CORRETOR") = IIf(CurrentQuery.FieldByName("CORRETOR").IsNull, "-1",CurrentQuery.FieldByName("CORRETOR").AsString)
	UserVar("COMPETENCIA") = CurrentQuery.FieldByName("COMPETENCIA").AsString


    If WebMode Then
		Dim sx As CSServerExec
		Set sx = NewServerExec

		Dim container As CSDContainer
		Set container = NewContainer

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		sx.Description = "Processo - COM001 - Demonstrativo de Faturamento de Comissão."
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SessionVar("codigo") = "COM001"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Prestadores"))
		sx.SetContainer(container)
		sx.Execute

		Set sx = Nothing
		Set container = Nothing
	 End If
End Sub
