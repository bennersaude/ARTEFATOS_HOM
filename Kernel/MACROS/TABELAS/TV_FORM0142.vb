'HASH: 16495CDC23CF4484C2BBB5CE8F258F2A
 '#Uses "*bsShowMessage"
Dim handleRotinaDoc As Integer

Public Sub TABLE_AfterInsert()
	Dim qSql As BPesquisa
	Set qSql = NewQuery

	handleRotinaDoc = RecordHandleOfTable("SFN_ROTINADOC")

	qSql.Add("SELECT B.DESCRICAO,                ")
	qSql.Add("       A.DATA,                     ")
	qSql.Add("       A.COMPETENCIADEMONSTRATIVO  ")
	qSql.Add("  FROM SFN_ROTINADOC     A,        ")
	qSql.Add("       SFN_TIPODOCUMENTO B         ")
	qSql.Add("  WHERE A.HANDLE = :PHANDLE        ")
	qSql.Add("  AND B.HANDLE = A.TIPODOCUMENTO   ")
	qSql.ParamByName("PHANDLE").AsInteger = handleRotinaDoc
	qSql.Active = True

	CurrentQuery.FieldByName("TIPODOCUMENTO").AsString          = qSql.FieldByName("DESCRICAO").AsString
	CurrentQuery.FieldByName("DATADOCUMENTO").AsDateTime        = qSql.FieldByName("DATA").AsDateTime
	CurrentQuery.FieldByName("COMPETENCIA").AsString            = qSql.FieldByName("COMPETENCIADEMONSTRATIVO").AsString
	CurrentQuery.FieldByName("DATADOCUMENTODESTINO").AsDateTime = CurrentSystem.ServerDate
	CurrentQuery.FieldByName("COMPETENCIADESTINO").AsDateTime   = CurrentSystem.ServerDate

	Set qSql = Nothing
End Sub

Public Sub TABLE_AfterPost()

	Dim bs As CSBusinessComponent

	Set bs = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnRotinaDocBLL, Benner.Saude.Financeiro.Business") ' formato: [namespace.classe], [assembly]

	bs.ClearParameters
	bs.AddParameter(pdtInteger, handleRotinaDoc)
	bs.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATADOCUMENTODESTINO").AsDateTime)
	bs.AddParameter(pdtDateTime, CurrentQuery.FieldByName("COMPETENCIADESTINO").AsDateTime)
	bs.Execute("Duplicar")
	Set bs = Nothing

	bsshowMessage("Rotina copiada com sucesso para a competência: "  & FormatDateTime2("MM/YYYY",CurrentQuery.FieldByName("COMPETENCIADESTINO").AsDateTime) ,"I")

End Sub

Public Sub TABLE_AfterScroll()
	TIPODOCUMENTO.ReadOnly = True
	DATADOCUMENTO.ReadOnly = True
	COMPETENCIA.ReadOnly = True
End Sub
