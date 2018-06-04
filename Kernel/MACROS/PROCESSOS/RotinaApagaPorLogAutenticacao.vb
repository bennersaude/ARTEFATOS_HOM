'HASH: 9626EDA0DA6FA1818647D44258083FFF

Public Sub Main

	Dim Sql As Object
	Dim periodo_data As Date
  	Set Sql = NewQuery

  	Dim Del As Object
  	Set Del = NewQuery

	If (SessionVar("PERIODO_MES") <>"") Then
		vsPERIODO_MES = SessionVar("PERIODO_MES")

		periodo_data = DateAdd("m", -vsperiodo_mes, ServerDate)

		Sql.Add("SELECT * ")
  		Sql.Add("  FROM POR_LOGAUTENTICACAO")
  		Sql.Add(" WHERE DATAHORA <= :data")
  		Sql.ParamByName("data").AsDateTime = periodo_data
  		Sql.Active = True

  		While Not Sql.EOF
    		Del.Clear
    		Del.Add("DELETE FROM POR_LOGAUTENTICACAO WHERE HANDLE = " + Sql.FieldByName("HANDLE").AsString)
    		Del.ExecSQL

    		Sql.Next
  		Wend

	Else
		Err.Raise -999,"Erro","O parâmetro PERIODO_MES deve ser definido!"
	End If

  	Set Sql = Nothing
  	Set Del = Nothing

End Sub
