'HASH: D1A473CB053E52E6BFB8809447715233

'macro sam_geradordv

Public Sub TABLE_AfterPost()
  Dim SQL As Object
  Dim SQLINSERT As Object
  Dim SQLMAX As Object
  Dim SQLDELETE As Object

  Dim vCont As Integer
  Dim vDigitos As Integer
  Dim vMax As Integer


  Set SQL = NewQuery
  Set SQLINSERT = NewQuery
  Set SQLMAX = NewQuery
  Set SQLDELETE = NewQuery

  SQL.Add("SELECT DIGITOS FROM SAM_GERADORDV WHERE HANDLE = :H")
  SQL.ParamByName("H").AsInteger = RecordHandleOfTable("SAM_GERADORDV")
  SQL.Active = True

  vDigitos = SQL.FieldByName("DIGITOS").AsInteger

  SQLINSERT.Add("INSERT INTO SAM_GERADORDV_PESO (HANDLE, NUMERO, GERADORDV) VALUES (:HANDLE, :NUMERO, :GERADORDV)")

  SQLMAX.Add("SELECT MAX(NUMERO) MAXIMO FROM SAM_GERADORDV_PESO WHERE GERADORDV =:GERADORDV")
  SQLMAX.ParamByName("GERADORDV").AsInteger = RecordHandleOfTable("SAM_GERADORDV")
  SQLMAX.Active = True

  SQLDELETE.Add("DELETE FROM SAM_GERADORDV_PESO WHERE NUMERO >:NUM AND GERADORDV =:GERADORDV")

  vMax = SQLMAX.FieldByName("MAXIMO").AsInteger

  If (SQLMAX.FieldByName("MAXIMO").IsNull) Then
    vCont = 1

  Else
    If vMax < vDigitos Then
      vCont = vMax + 1
    Else
      SQLDELETE.ParamByName("NUM").AsInteger = vDigitos
      SQLDELETE.ParamByName("GeradorDv").AsInteger = RecordHandleOfTable("SAM_GERADORDV")
      SQLDELETE.ExecSQL
      vCont = vDigitos + 1
    End If
  End If
  While vCont <= vDigitos
    SQLINSERT.ParamByName("NUMERO").AsInteger = vCont
    SQLINSERT.ParamByName("HANDLE").AsInteger = NewHandle("SAM_GERADORDV_PESO")
    SQLINSERT.ParamByName("GERADORDV").AsInteger = RecordHandleOfTable("SAM_GERADORDV")
    SQLINSERT.ExecSQL
    vCont = vCont + 1
  Wend

End Sub

