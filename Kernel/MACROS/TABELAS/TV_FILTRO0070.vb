'HASH: 2C903079559975E27302EB3932BDA457
'#Uses "*Biblioteca"

Public Sub TABLE_AfterInsert()
  Dim qUsuario As Object
  Set qUsuario = NewQuery

  If UserVar("FILTRO_TV_FILTRO070") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FILTRO070"),CurrentQuery.TQuery)
  End If


  With qUsuario
    .Active = False
    .Clear
    .Add("SELECT PRESTADOR")
    .Add("  FROM Z_GRUPOUSUARIOS_PRESTADOR")
    .Add(" WHERE USUARIO = :USUARIO")
    .ParamByName("USUARIO").AsInteger = CurrentUser
    .Active = True
  End With

  If Not(qUsuario.FieldByName("PRESTADOR").IsNull) Then
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = qUsuario.FieldByName("PRESTADOR").AsInteger
  End If

  Set qUsuario = Nothing
End Sub


Public Sub TABLE_AfterPost()
	UserVar("FILTRO_TV_FILTRO070") = DatasetToXML(CurrentQuery.TQuery,"")


If WebMode Then
		UserVar("TIPOFATURAMENTO")  = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString
		UserVar("COMPETENCIA")      = CurrentQuery.FieldByName("COMPETENCIA").AsString
		UserVar("ROTINAFINANCEIRA") = CurrentQuery.FieldByName("ROTINAFINANCEIRA").AsString
		UserVar("PRESTADOR")        = CurrentQuery.FieldByName("PRESTADOR").AsString
		UserVar("ORDENARPORGUIA")   = CurrentQuery.FieldByName("ORDENARPORGUIA").AsString
		UserVar("ORDENARPORNOME")   = CurrentQuery.FieldByName("ORDENARPORNOME").AsString
		UserVar("EMITIRRESUMO")     = CurrentQuery.FieldByName("EMITIRRESUMO").AsString

		Dim container As CSDContainer
		Set container = NewContainer

		container.GetFieldsFromQuery(CurrentQuery.TQuery)
		container.LoadAllFromQuery(CurrentQuery.TQuery)

		Dim sx As CSServerExec
		Set sx = NewServerExec
		sx.Description = "Relatório - DEM-PG1-Demonstrativo de Pagamento Prestador"
		sx.Process = RetornaHandleProcesso("RELATORIO_STIMULSOFT")
		sx.SetContainer(container)
		sx.SessionVar("codigo") = "DEMPG1"


		sx.SessionVar("modulo") = CStr(RetornaHandleModulo("Controle Financeiro"))
		sx.Execute
		Set sx = Nothing

		Set container = Nothing

	End If

End Sub





