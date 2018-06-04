'HASH: 9EF193A0D1F6FBA6B46735E27C73F837

'#Uses "*bsShowMessage"

'MACRO TABELA: SAM_PRESTADOR_PROC_REGEXC_REDE

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
    Else
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

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim SQL As Object
  Dim qPrest As Object
  Dim REDES As Object
  Dim vRede As Long

  If WebMode Then
    Set qPrest = NewQuery

  	qPrest.Add("SELECT A3.PRESTADOR")
  	qPrest.Add("  FROM SAM_PRESTADOR_PROC_REGEXC A1")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC_CREDEN A2 ON A1.PRESTADORPROCESSO = A2.HANDLE")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC A3 ON A2.PRESTADORPROCESSO = A3.HANDLE")
  	qPrest.Add(" WHERE A1.HANDLE = :HANDLE")
    qPrest.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCREGEXC").Value

    qPrest.Active = True
  End If

  Set SQL = NewQuery
  SQL.Add("SELECT REDERESTRITA, PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  SQL.Add("AND DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")

  If VisibleMode Then
  	SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  Else
	SQL.ParamByName("PRESTADOR").Value = qPrest.FieldByName("PRESTADOR").Value
  End If

  SQL.ParamByName("DATA").Value = ServerDate
  SQL.Active = True

  vCondicao = ""
  If WebMode Then
    vCondicao = vCondicao + "A.HANDLE "
  Else
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
      Else
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

Public Sub TABLE_UpdateRequired()
Dim SQL1 As Object
  Dim qPrest As Object
  Dim REDES As Object
  Dim vRede As Long

  If WebMode Then
    Set qPrest = NewQuery

  	qPrest.Add("SELECT A3.PRESTADOR")
  	qPrest.Add("  FROM SAM_PRESTADOR_PROC_REGEXC A1")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC_CREDEN A2 ON A1.PRESTADORPROCESSO = A2.HANDLE")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC A3 ON A2.PRESTADORPROCESSO = A3.HANDLE")
  	qPrest.Add(" WHERE A1.HANDLE = :HANDLE")
    qPrest.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCREGEXC").Value

    qPrest.Active = True
  End If


  Set SQL1 = NewQuery
  SQL1.Add("SELECT REDERESTRITA, PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  SQL1.Add("AND DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")
  If VisibleMode Then
  	SQL1.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  Else
	SQL1.ParamByName("PRESTADOR").Value = qPrest.FieldByName("PRESTADOR").Value
  End If
  SQL1.ParamByName("DATA").Value = ServerDate
  SQL1.Active = True

  vCondicao = ""
  If WebMode Then
    vCondicao = vCondicao + "A.HANDLE "
  Else
    vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
  End If
  vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = " + SQL1.FieldByName("REDERESTRITA").AsInteger + ")"

  Set REDES = NewQuery
  REDES.Add("SELECT REDERESTRITA, REDERESTRITACONTIDA FROM SAM_REDERESTRITACONTIDA WHERE REDERESTRITA = :REDERESTRITA")
  REDES.ParamByName("REDERESTRITA").Value = SQL1.FieldByName("REDERESTRITA").AsInteger
  REDES.Active = True

  While Not SQL1.EOF
    vRede = SQL1.FieldByName("REDERESTRITA").AsInteger
    Recursividade(vRede)
    SQL1.Next
    If Not SQL1.EOF Then
      vCondicao = vCondicao + " OR ("
      If WebMode Then
        vCondicao = vCondicao + "A.HANDLE "
      Else
        vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
      End If
      vCondicao = vCondicao + "IN (SELECT REDERESTRITA FROM SAM_REDERESTRITA_PRESTADOR WHERE REDERESTRITA = " + SQL1.FieldByName("REDERESTRITA").AsInteger + ")"
      vCondicao = vCondicao + "    )"
    End If
  Wend

  Set REDES = Nothing

  If WebMode Then
  	REDERESTRITA.WebLocalWhere = vCondicao
  ElseIf VisibleMode Then
  	REDERESTRITA.LocalWhere = vCondicao
  End If


  Set SQL1 = Nothing

 Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_REGEXC WHERE HANDLE = :HANDLE")
  If VisibleMode Then
    SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_REGEXC")
  Else
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADORPROCREGEXC").Value
  End If
  SQL.Active = True
  If SQL.FieldByName("OPERACAO").AsString = "E" Then
    CanContinue = False
    bsShowMessage("O tipo de operação não permite inserir registros nesta carga !", "E")
  End If
  Set SQL = Nothing
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim FILIADOS As Object
  Dim linha As String

  Set FILIADOS = NewQuery
  FILIADOS.Add("SELECT GERARPARAFILIADOS FROM SAM_PRESTADOR_PROC_REGEXC WHERE HANDLE=:HANDLE")
  FILIADOS.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_REGEXC")
  FILIADOS.Active = True

  If FILIADOS.FieldByName("GERARPARAFILIADOS").Value = "S" Then
    linha = "Não será permitido a inclusão de redes!!!" + Chr(10)
    linha = linha + "Motivo: Prestados filiados já foram selecionados."
    bsShowMessage(linha, "E")
    CanContinue = False
  End If

  Set FILIADOS = Nothing
  '--------------------------------------------------------
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT *                                          ")
  SQL.Add("  FROM SAM_PRESTADOR_PROC_REGEXC_REDE             ")
  SQL.Add(" WHERE HANDLE <> :HANDLE                          ")
  SQL.Add("   AND REDERESTRITA = :REDERESTRITA               ")
  SQL.Add("   AND PRESTADORPROCREGEXC = :PRESTADORPROCREGEXC ")

  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.ParamByName("REDERESTRITA").Value = CurrentQuery.FieldByName("REDERESTRITA").Value
  SQL.ParamByName("PRESTADORPROCREGEXC").Value = CurrentQuery.FieldByName("PRESTADORPROCREGEXC").Value
  SQL.Active = True
  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Esta rede já está cadastrada !", "E")
  End If
  Set SQL = Nothing
  '--------------------------------------------------------
End Sub

