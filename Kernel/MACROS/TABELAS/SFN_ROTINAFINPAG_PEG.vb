'HASH: 3E8DC705B148DC9E217E33B1EB46E9B3
Public Function RotinaDeReapresentacao
  Dim SQL As BPesquisa
  Set SQL = NewQuery
  SQL.Add("SELECT B.EHREAPRESENTACAO ")
  SQL.Add("  FROM SIS_TIPOFATURAMENTO A ")
  SQL.Add("  JOIN SFN_ROTINAFIN       B ON A.Handle = B.TIPOFATURAMENTO ")
  SQL.Add(" WHERE B.Handle = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Active = True

  RotinaDeReapresentacao = SQL.FieldByName("EHREAPRESENTACAO").AsString = "S"

  Set SQL = Nothing

End Function

Public Function RotinaDeCredenciamento
  Dim SQL As BPesquisa
  Set SQL = NewQuery
  SQL.Add("SELECT A.CODIGO ")
  SQL.Add("  FROM SIS_TIPOFATURAMENTO A ")
  SQL.Add("  JOIN SFN_ROTINAFIN       B ON A.Handle = B.TIPOFATURAMENTO ")
  SQL.Add(" WHERE B.Handle = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Active = True

  RotinaDeCredenciamento = 1

  If SQL.FieldByName("CODIGO").AsInteger = 310 Then
    RotinaDeCredenciamento = 2
  End If

  Set SQL = Nothing

End Function

Public Function RetornaDataPagamento
	Dim SQL As BPesquisa
	Set SQL = NewQuery

	SQL.Active = False
	SQL.Clear
	SQL.Add("SELECT DATAPAGAMENTO FROM SFN_ROTINAFINPAG WHERE HANDLE = :HANDLE")
	SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINAFINPAG")
	SQL.Active = True

	RetornaDataPagamento = SQL.FieldByName("DATAPAGAMENTO").AsDateTime

	Set SQL = Nothing
End Function

Public Function PermiteFaturarReapresentadoJuntoNormal
  Dim SQL As BPesquisa
  Set SQL = NewQuery
  SQL.Add("SELECT P.REAPRESENTADOJUNTONORMAL ")
  SQL.Add("  FROM SAM_PARAMETROSPROCCONTAS P ")
  SQL.Active = True

  PermiteFaturarReapresentadoJuntoNormal = SQL.FieldByName("REAPRESENTADOJUNTONORMAL").AsString = "S"

  Set SQL = Nothing

End Function


Public Function retornaFiltroPEG
  Dim vAlias As String

  vAlias = "SAM_PEG"
  If WebMode Then
    vAlias = "A"
  End If

  Dim rp As Integer
  Dim vEhReapresentacao As Boolean
  Dim vEhControlePagamento As Boolean
  Dim vPermiteFaturarReapresentadoJuntoNormal As Boolean
  Dim datapag As Date

  vEhReapresentacao = RotinaDeReapresentacao
  vEhControlePagamento = RotinaDeControlePagamento
  vPermiteFaturarReapresentadoJuntoNormal = PermiteFaturarReapresentadoJuntoNormal

  rp = RotinaDeCredenciamento

  datapag = RetornaDataPagamento

  Dim STRX As String
  If InStr(SQLServer, "DB2") > 0 Then
    STRX = " AND timestamp_iso(date(" + vAlias + ".DATAPAGAMENTO)) = "
  ElseIf InStr(SQLServer, "ORACLE") > 0 Then
    STRX = " AND trunc(" + vAlias + ".DATAPAGAMENTO) = "
  ElseIf InStr(SQLServer, "MSSQL") > 0 Then
    STRX = " AND convert(datetime, cast(floor(convert(float, " + vAlias + ".DATAPAGAMENTO)) as int)) = "
  Else
    STRX = " AND CONVERT(DATETIME , CAST(" + vAlias + ".DATAPAGAMENTO  AS DATE), 103) ="
  End If

  retornaFiltroPEG = "     " + vAlias + ".SITUACAO = 3 " + _
					 " AND " + vAlias + ".TABREGIMEPGTO = " + Str(rp) + " " + STRX + SQLDate(datapag) + _
					 " AND " + vAlias + ".COMPETENCIA = (SELECT COMPETENCIA " + _
					 "                                     FROM SFN_ROTINAFINPAG " + _
					 "                                    WHERE HANDLE = " + CStr(RecordHandleOfTable("SFN_ROTINAFINPAG"))+ ")" + _
					 " AND " + vAlias + ".HANDLE NOT IN (SELECT PEG FROM SFN_ROTINAFINPAG_PEG WHERE ROTINAFINPAG = " + CStr(RecordHandleOfTable("SFN_ROTINAFINPAG"))+ ")"

  If vEhReapresentacao Then
    retornaFiltroPEG = retornaFiltroPEG + " AND " + vAlias + ".PEGORIGINAL IS NOT NULL"
  ElseIf (Not vEhControlePagamento) And (Not vPermiteFaturarReapresentadoJuntoNormal) Then
    retornaFiltroPEG = retornaFiltroPEG + " AND " + vAlias + ".PEGORIGINAL IS NULL"
  End If

End Function

Public Function RotinaDeControlePagamento
  Dim SQL As BPesquisa
  Set SQL = NewQuery
  SQL.Add("SELECT B.CONTROLEPAGAMENTO ")
  SQL.Add("  FROM SIS_TIPOFATURAMENTO A ")
  SQL.Add("  JOIN SFN_ROTINAFIN       B ON A.Handle = B.TIPOFATURAMENTO ")
  SQL.Add(" WHERE B.Handle = :HANDLE")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Active = True

  RotinaDeControlePagamento = SQL.FieldByName("CONTROLEPAGAMENTO").AsString = "S"

  Set SQL = Nothing

End Function


Public Sub PEG_OnPopup(ShowPopup As Boolean)

  PEG.LocalWhere = retornaFiltroPEG

End Sub

Public Sub TABLE_AfterScroll()
  If WebMode Then
    PEG.WebLocalWhere = retornaFiltroPEG
  Else
    PEG.LocalWhere = retornaFiltroPEG
  End If
End Sub
