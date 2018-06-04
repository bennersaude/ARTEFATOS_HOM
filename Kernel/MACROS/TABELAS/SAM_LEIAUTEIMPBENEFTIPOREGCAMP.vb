'HASH: 2D22D5B6F39D8879D5911CCD0C56A2E4

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If WebMode Then
	CAMPO.WebLocalWhere = "A.HANDLE IN (SELECT C.HANDLE FROM SIS_CAMPOS C " & _
  											"JOIN SAM_LEIAUTEIMPBENEFTIPOREG L On L.HANDLE = @TIPOREGISTRO " & _
  											"JOIN SAM_LEIAUTEIMPBENEF LE ON LE.HANDLE = L.LEIAUTEIMPBENEF " & _
  											"WHERE (C.TIPOREGISTRO = L.CODIGO OR C.PERMITIRTODOSTIPOREG = 'S') " & _
  											"  AND (C.LEIAUTEDE = 'A' OR C.LEIAUTEDE = LE.LEIAUTEDE)" & _
  											") "


  ElseIf VisibleMode Then
	CAMPO.LocalWhere = "SIS_CAMPOS.HANDLE IN (SELECT C.HANDLE FROM SIS_CAMPOS C " & _
  											"JOIN SAM_LEIAUTEIMPBENEFTIPOREG L On L.HANDLE = @TIPOREGISTRO " & _
  											"JOIN SAM_LEIAUTEIMPBENEF LE ON LE.HANDLE = L.LEIAUTEIMPBENEF " & _
  											"WHERE (C.TIPOREGISTRO = L.CODIGO OR C.PERMITIRTODOSTIPOREG = 'S') " & _
  											"  AND (C.LEIAUTEDE = 'A' OR C.LEIAUTEDE = LE.LEIAUTEDE)" & _
  											") "
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If WebMode Then
	CAMPO.WebLocalWhere = "A.HANDLE IN (SELECT C.HANDLE FROM SIS_CAMPOS C " & _
  											"JOIN SAM_LEIAUTEIMPBENEFTIPOREG L On L.HANDLE = @TIPOREGISTRO " & _
  											"JOIN SAM_LEIAUTEIMPBENEF LE ON LE.HANDLE = L.LEIAUTEIMPBENEF " & _
  											"WHERE (C.TIPOREGISTRO = L.CODIGO OR C.PERMITIRTODOSTIPOREG = 'S') " & _
  											"  AND (C.LEIAUTEDE = 'A' OR C.LEIAUTEDE = LE.LEIAUTEDE)" & _
  											") "


  ElseIf VisibleMode Then
	CAMPO.LocalWhere = "SIS_CAMPOS.HANDLE IN (SELECT C.HANDLE FROM SIS_CAMPOS C " & _
  											"JOIN SAM_LEIAUTEIMPBENEFTIPOREG L On L.HANDLE = @TIPOREGISTRO " & _
  											"JOIN SAM_LEIAUTEIMPBENEF LE ON LE.HANDLE = L.LEIAUTEIMPBENEF " & _
  											"WHERE (C.TIPOREGISTRO = L.CODIGO OR C.PERMITIRTODOSTIPOREG = 'S') " & _
  											"  AND (C.LEIAUTEDE = 'A' OR C.LEIAUTEDE = LE.LEIAUTEDE)" & _
  											") "
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 2 Then 'alteração
    Dim qUpd As Object
    Set qUpd = NewQuery

    Dim qLeiaute As Object
    Set qLeiaute = NewQuery

    qLeiaute.Active = False
    qLeiaute.Clear
    qLeiaute.Add("SELECT LEIAUTEIMPBENEF")
    qLeiaute.Add("  FROM SAM_LEIAUTEIMPBENEFTIPOREG ")
    qLeiaute.Add(" WHERE HANDLE = :TIPOREGISTRO")
    qLeiaute.ParamByName("TIPOREGISTRO").AsInteger = CurrentQuery.FieldByName("TIPOREGISTRO").AsInteger
    qLeiaute.Active = True


    qUpd.Active = False
    qUpd.Clear
    qUpd.Add("UPDATE SAM_LEIAUTEIMPBENEF                  ")
    qUpd.Add("   SET DATAALTERACAO = :DATAALTERACAO,      ")
    qUpd.Add("       USUARIOALTERACAO = :USUARIOALTERACAO ")
    qUpd.Add(" WHERE HANDLE = :HLEIAUTE")

    qUpd.ParamByName("DATAALTERACAO").AsDateTime = ServerDate
    qUpd.ParamByName("USUARIOALTERACAO").AsInteger = CurrentUser
    qUpd.ParamByName("HLEIAUTE").AsInteger = qLeiaute.FieldByName("LEIAUTEIMPBENEF").AsInteger
    qUpd.ExecSQL


    Set qUpd = Nothing
    Set qLeiaute = Nothing

  End If

End Sub
