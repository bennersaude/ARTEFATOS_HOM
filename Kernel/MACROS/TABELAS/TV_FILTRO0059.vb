'HASH: 1F8BFEA8E5B044E0A0FB1F3DC1DDEC14
 

Public Sub TABLE_AfterInsert()
  Dim qUsuario As Object
  Set qUsuario = NewQuery

  If UserVar("FILTRO_TV_FILTRO059") <> "" Then
		XMLToDataset(UserVar("FILTRO_TV_FILTRO059"),CurrentQuery.TQuery)
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
	UserVar("FILTRO_TV_FILTRO059") = DatasetToXML(CurrentQuery.TQuery,"")
End Sub
