'HASH: 3399B71A499D4C7B9BAF9BBC5E17570F
'#uses "*Biblioteca"

Option Explicit

Public Sub TABLE_AfterInsert()
	If UserVar("TV_FILTROALTERACAOBENEFICIARIO") <> "" Then
		XMLToDataset(UserVar("TV_FILTROALTERACAOBENEFICIARIO"),CurrentQuery.TQuery)
	End If
End Sub

Public Sub TABLE_AfterPost()
	UserVar("TV_FILTROALTERACAOBENEFICIARIO") = DatasetToXML(CurrentQuery.TQuery,"")

	CurrentQuery.Edit
	If Not CurrentQuery.FieldByName("DATASOLICITACAO").IsNull Then
		CurrentQuery.FieldByName("DATASOLICITACAOADICIONAL").Value = CurrentQuery.FieldByName("DATASOLICITACAO").AsDateTime +1
	Else
		CurrentQuery.FieldByName("DATASOLICITACAO").Value = Null
	End If

	If WebMode Then
		Dim container As CSDContainer
		Set container = NewContainer

		Dim vFiltro As String

		vFiltro = " B.CONTRATO IN (" & ConvertPipeToVirgulaCampoFiltro(CurrentQuery.FieldByName("BENEFICIARIO").AsString) & ") "
		vFiltro = vFiltro & IIf(Not CurrentQuery.FieldByName("NUMEROFAMILIA").IsNull, " AND F.HANDLE IN (" & ConvertPipeToVirgulaCampoFiltro(CurrentQuery.FieldByName("NUMEROFAMILIA").AsString) & ") ","")
		vFiltro = vFiltro & IIf(Not CurrentQuery.FieldByName("SITUACAOREGISTRO").IsNull, " AND W.SITUACAO = (" & ConvertPipeToVirgulaCampoFiltro(CurrentQuery.FieldByName("SITUACAOREGISTRO").AsString) & ") ","")

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Alteração cadastral de beneficiário"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "BEN205"
		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Beneficiários"))
		sx.SessionVar("DefaultWhere") = vFiltro
		sx.Execute

		Set sx = Nothing
		Set container = Nothing

	End If
End Sub
