'HASH: D8754A30CE57B217CAB8D60B2ECEF934

'MACRO TABELA: SAM_PRESTADOR_REGRAREDE
'#Uses "*bsShowMessage"

Dim vCondicao As String


Sub Recursividade(pRede As Long)
  Dim CONTIDAS As Object
  Dim vRede As Long


  Set CONTIDAS = NewQuery

  CONTIDAS.Add("SELECT REDERESTRITA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITACONTIDA = :REDERESTRITA")
  CONTIDAS.ParamByName("REDERESTRITA").Value = pRede
  CONTIDAS.Active = True
  If Not CONTIDAS.EOF Then
    vCondicao = vCondicao + " OR ("
    If WebMode Then
    	vCondicao = vCondicao + "A.HANDLE "
    ElseIf VisibleMode Then
		vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
    End If
    vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITACONTIDA = " + CStr(pRede) + ")"
    vCondicao = vCondicao + "    )"

    While Not CONTIDAS.EOF
      vRede = CONTIDAS.FieldByName("REDERESTRITA").AsInteger
      Recursividade(vRede)
      CONTIDAS.Next
    Wend
  End If
  Set CONTIDAS = Nothing

End Sub


Public Sub TABLE_AfterInsert()

  Dim SQL As Object
  Dim REDES As Object
  Dim vRede As Long


  Set SQL = NewQuery
  SQL.Add("SELECT REDERESTRITA, PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  SQL.Add("AND DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")
  SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  SQL.ParamByName("DATA").Value = ServerDate
  SQL.Active = True

  vCondicao = ""
  If WebMode Then
  	vCondicao = vCondicao + "A.HANDLE "
  ElseIf VisibleMode Then
  	vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
  End If
  vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = " + SQL.FieldByName("REDERESTRITA").AsInteger + ")"

  Set REDES = NewQuery
  REDES.Add("SELECT REDERESTRITA, REDERESTRITACONTIDA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDERESTRITA")
  REDES.ParamByName("REDERESTRITA").Value = SQL.FieldByName("REDERESTRITA").AsInteger
  REDES.Active = True


  While Not SQL.EOF
    vRede = SQL.FieldByName("REDERESTRITA").AsInteger
    Recursividade(vRede)
    SQL.Next
    If Not SQL.EOF Then
      vCondicao = vCondicao + " OR ("
      If WebMode Then
      	vCondicao = vCondicao + "A.HANDLE "
      ElseIf VisibleMode Then
      	vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
      End If
      vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = " + SQL.FieldByName("REDERESTRITA").AsInteger + ")"
      vCondicao = vCondicao + "    )"
    End If
  Wend

  Set REDES = Nothing

  If WebMode Then
  	REDERESTRITA.WebLocalWhere = vCondicao
  ElseIf VisibleMode Then
    REDERESTRITA.LocalWhere = vCondicao
  End If
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Dim REDES As Object
  Dim vRede As Long


  Set SQL = NewQuery
  SQL.Add("SELECT REDERESTRITA, PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  SQL.Add("AND DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")
  SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  SQL.ParamByName("DATA").Value = ServerDate
  SQL.Active = True

  vCondicao = ""
  If WebMode Then
  	vCondicao = vCondicao + "A.HANDLE "
  ElseIf VisibleMode Then
  	vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
  End If
  vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = " + SQL.FieldByName("REDERESTRITA").AsInteger + ")"

  Set REDES = NewQuery
  REDES.Add("SELECT REDERESTRITA, REDERESTRITACONTIDA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDERESTRITA")
  REDES.ParamByName("REDERESTRITA").Value = SQL.FieldByName("REDERESTRITA").AsInteger
  REDES.Active = True


  While Not SQL.EOF
    vRede = SQL.FieldByName("REDERESTRITA").AsInteger
    Recursividade(vRede)
    SQL.Next
    If Not SQL.EOF Then
      vCondicao = vCondicao + " OR ("
      If WebMode Then
      	vCondicao = vCondicao + "A.HANDLE "
      ElseIf VisibleMode Then
      	vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
      End If
      vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = " + SQL.FieldByName("REDERESTRITA").AsInteger + ")"
      vCondicao = vCondicao + "    )"
    End If
  Wend

  Set REDES = Nothing

  If WebMode Then
  	REDERESTRITA.WebLocalWhere = vCondicao
  ElseIf VisibleMode Then
    REDERESTRITA.LocalWhere = vCondicao
  End If
  Set SQL = Nothing
End Sub
