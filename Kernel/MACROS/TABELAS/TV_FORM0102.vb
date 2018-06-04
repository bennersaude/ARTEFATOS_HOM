'HASH: 6D7818EBBC7CC3DC8DCF56FD70628A2F
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	Dim qSql As BPesquisa
	Set qSql = NewQuery
	qSql.Add("SELECT A.DESCRICAO,                ")
	qSql.Add("       B.COMPETENCIA,              ")
	qSql.Add("       C.SEQUENCIA,                ")
	qSql.Add("       C.DATAROTINA                ")
	qSql.Add("  FROM SIS_TIPOFATURAMENTO A,      ")
	qSql.Add("       SFN_COMPETFIN B,            ")
	qSql.Add("       SFN_ROTINAFIN C,            ")
	qSql.Add("       SFN_ROTINAFINFAT D          ")
	qSql.Add("  WHERE D.HANDLE = :PHANDLE        ")
	qSql.Add("  AND C.HANDLE = D.ROTINAFIN       ")
	qSql.Add("  AND B.HANDLE = C.COMPETFIN       ")
	qSql.Add("  AND A.HANDLE = B.TIPOFATURAMENTO ")
	qSql.ParamByName("PHANDLE").AsInteger = RecordHandleOfTable("SFN_ROTINAFINFAT")
	qSql.Active = True

	CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString      = qSql.FieldByName("DESCRICAO").AsString
	CurrentQuery.FieldByName("COMPETENCIA").AsDateTime        = qSql.FieldByName("COMPETENCIA").AsDateTime
	CurrentQuery.FieldByName("SEQUENCIA").AsString            = qSql.FieldByName("SEQUENCIA").AsString
	CurrentQuery.FieldByName("DATAROTINA").AsDateTime         = qSql.FieldByName("DATAROTINA").AsDateTime
	CurrentQuery.FieldByName("COMPETENCIADESTINO").AsDateTime = qSql.FieldByName("COMPETENCIA").AsDateTime

	Set qSql = Nothing
End Sub

Public Sub TABLE_AfterPost()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SAMFaturamento.Faturamento")

  Obj.Duplicar(CurrentSystem, RecordHandleOfTable("SFN_ROTINAFINFAT"))

  Set Obj = Nothing

  bsshowMessage("Rotina copiada com sucesso para a competência: "  & FormatDateTime2("MM/YYYY",CurrentQuery.FieldByName("COMPETENCIADESTINO").AsDateTime) ,"I")

End Sub

Public Sub TABLE_AfterScroll()
	TIPOFATURAMENTO.ReadOnly = True
	COMPETENCIA.ReadOnly = True
	SEQUENCIA.ReadOnly = True
	DATAROTINA.ReadOnly = True
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	SessionVar("COMPETENCIA_DUPLICAR") = FormatDateTime2("MM/YYYY",CurrentQuery.FieldByName("COMPETENCIADESTINO").AsDateTime)
End Sub
