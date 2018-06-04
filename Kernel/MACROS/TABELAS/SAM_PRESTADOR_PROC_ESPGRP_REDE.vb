'HASH: 3B80C9F0E5609E193438A085A003B4F9

'MACRO TABELA: SAM_PRESTADOR_PROC_ESPGRP_REDE

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
    vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
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
  Dim qPrest As Object
  Dim vRede As Long

  If WebMode Then
    Set qPrest = NewQuery
  	qPrest.Add("SELECT A4.PRESTADOR")
  	qPrest.Add("  FROM SAM_PRESTADOR_PROC_ESPEC_GRP A1")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC_ESPEC A2 ON A1.PRESTADORPROCESSO = A2.HANDLE")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC_CREDEN A3 ON A2.PRESTADORPROCESSO = A3.HANDLE")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC A4 ON A3.PRESTADORPROCESSO = A4.HANDLE")
  	qPrest.Add(" WHERE A1.HANDLE = :HANDLE")
	qPrest.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCESSOESPECIALIDADEGRUPO")
	qPrest.Active = True
  End If

  Set SQL = NewQuery
  SQL.Add("SELECT REDERESTRITA, PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  SQL.Add("AND DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")
  If VisibleMode Then
  	SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  Else
	SQL.ParamByName("PRESTADOR").Value = qPrest.FieldByName("PRESTADOR")
  End If
  SQL.ParamByName("DATA").Value = ServerDate
  SQL.Active = True

  vCondicao = ""
  vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
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
      vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
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
  Dim qPrest As Object
  Dim REDES As Object
  Dim vRede As Long

  If WebMode Then
	Set qPrest = NewQuery

  	qPrest.Add("SELECT A4.PRESTADOR")
  	qPrest.Add("  FROM SAM_PRESTADOR_PROC_ESPEC_GRP A1")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC_ESPEC A2 ON A1.PRESTADORPROCESSO = A2.HANDLE")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC_CREDEN A3 ON A2.PRESTADORPROCESSO = A3.HANDLE")
  	qPrest.Add("  JOIN SAM_PRESTADOR_PROC A4 ON A3.PRESTADORPROCESSO = A4.HANDLE")
  	qPrest.Add(" WHERE A1.HANDLE = :HANDLE")
	qPrest.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCESSOESPECIALIDADEGRUPO")
	qPrest.Active = True
  End If


  Set SQL = NewQuery
  SQL.Add("SELECT REDERESTRITA, PRESTADOR FROM SAM_REDERESTRITA_PRESTADOR WHERE PRESTADOR = :PRESTADOR")
  SQL.Add("AND DATAINICIAL <= :DATA AND (DATAFINAL >= :DATA OR DATAFINAL IS NULL)")
  If VisibleMode Then
  	SQL.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  Else
	SQL.ParamByName("PRESTADOR").Value = qPrest.FieldByName("PRESTADOR")
  End If
  SQL.ParamByName("DATA").Value = ServerDate
  SQL.Active = True

  vCondicao = ""
  vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
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
      vCondicao = vCondicao + "SAM_REDERESTRITA.HANDLE "
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim FILIADOS As Object
  Dim linha As String

  Set FILIADOS = NewQuery
  FILIADOS.Add("SELECT GERARPARAFILIADOS FROM SAM_PRESTADOR_PROC_ESPEC WHERE HANDLE=:HANDLE")
  If VisibleMode Then
  	FILIADOS.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_ESPEC")
  Else
  	FILIADOS.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PROCESSOESPECIALIDADEGRUPO")
  End If
  FILIADOS.Active = True

  If FILIADOS.FieldByName("GERARPARAFILIADOS").Value = "S" Then
    linha = "Não será permitido a inclusão de redes!!!" + Chr(10)
    linha = linha + "Motivo: Prestados filiados já foram selecionados."
    bsShowMessage(linha, "E")
    CanContinue = False
  End If

  Set FILIADOS = Nothing

End Sub

