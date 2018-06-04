'HASH: 259047BB1848786BBE7F8FD88FF35034
 
Public Sub TABLE_AfterInsert()
  Dim qUsuario As Object
  Set qUsuario = NewQuery

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
