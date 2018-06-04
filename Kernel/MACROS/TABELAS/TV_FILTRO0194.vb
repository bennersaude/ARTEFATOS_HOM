'HASH: B92D11E50CEF23640EC2C1F3E478A5AF
 
'#uses "*Biblioteca"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("FILTRO_TV_FORM_COM002") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FORM_COM002"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FORM_COM002") = DatasetToXML(CurrentQuery.TQuery,"")

	UserVar("CORRETOR") = IIf(CurrentQuery.FieldByName("CORRETOR").IsNull, "-1",CurrentQuery.FieldByName("CORRETOR").AsString)
	UserVar("ROTAPROPRIACAO") = CurrentQuery.FieldByName("ROTAPROPRIACAO").AsString


    If WebMode Then
		Dim sx As CSServerExec
		Set sx = NewServerExec

		Dim container As CSDContainer
		Set container = NewContainer

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		sx.Description = "Processo - COM001 - Demonstrativo de Pagamento de Comissão."
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SessionVar("codigo") = "COM002"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Prestadores"))
		sx.SetContainer(container)
		sx.Execute

		Set sx = Nothing
		Set container = Nothing
	 End If
End Sub
