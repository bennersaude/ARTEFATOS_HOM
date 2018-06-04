'HASH: 1EE1FFCADD61C3322867915F23C3D047
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	Dim qDocumento As BPesquisa
	Set qDocumento = NewQuery

	qDocumento.Clear
	qDocumento.Add("SELECT * FROM SFN_DOCUMENTO WHERE HANDLE = :HANDLE")
	qDocumento.ParamByName("HANDLE").AsInteger = CLng(SessionVar("WebHandleDocumento")) 'Variavel de sessão criada no afterscroll da tabela 'SFN_DOCUMENTO'
	qDocumento.Active = True

	CurrentQuery.FieldByName("VENCIMENTOANTERIOR").AsDateTime = qDocumento.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("NOVOVENCIMENTO").AsDateTime = qDocumento.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("VALOR").AsFloat = qDocumento.FieldByName("VALOR").AsFloat
	CurrentQuery.FieldByName("VALORCORRECAO").AsFloat = qDocumento.FieldByName("VALORCORRECAO").AsFloat
	CurrentQuery.FieldByName("VALORJUROS").AsFloat = qDocumento.FieldByName("VALORJURO").AsFloat
	CurrentQuery.FieldByName("VALORMULTA").AsFloat = qDocumento.FieldByName("VALORMULTA").AsFloat
	CurrentQuery.FieldByName("VALORDESCONTO").AsFloat = qDocumento.FieldByName("VALORDESCONTO").AsFloat

	CurrentQuery.FieldByName("VALORTOTAL").AsFloat = (CurrentQuery.FieldByName("VALOR").AsFloat + CurrentQuery.FieldByName("VALORJUROS").AsFloat + CurrentQuery.FieldByName("VALORMULTA").AsFloat + CurrentQuery.FieldByName("VALORCORRECAO").AsFloat) - CurrentQuery.FieldByName("VALORDESCONTO").AsFloat

	qDocumento.Active = False

	Set qDocumento = Nothing
End Sub

Public Sub TABLE_AfterPost()
	SessionVar("NovoVencimento") = FormatDateTime2("DD/MM/YYYY",CurrentQuery.FieldByName("NOVOVENCIMENTO").AsDateTime)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As BPesquisa
	Dim qDocumento As BPesquisa
	Set qDocumento = NewQuery

	qDocumento.Clear
	qDocumento.Add("SELECT * FROM SFN_DOCUMENTO WHERE HANDLE = :HANDLE")
	qDocumento.ParamByName("HANDLE").AsInteger = CLng(SessionVar("WebHandleDocumento"))
	qDocumento.Active = True

	If Not qDocumento.FieldByName("BAIXADATA").IsNull Then
		Err.Raise(vbsUserException,"","Documento já baixado")
		CanContinue = False
		Set qDocumento = Nothing
		Exit Sub
	End If

	Set SQL = NewQuery
	SQL.Clear
	SQL.Active = False
	SQL.Add("SELECT COUNT(1) QTDE")
  	SQL.Add("  FROM SFN_ROTINAARQUIVO_DOC RAD")
  	SQL.Add("  Join SFN_DOCUMENTO D     On (RAD.DOCUMENTO     = D.Handle)")
  	SQL.Add("  Join SFN_ROTINAARQUIVO R On (RAD.ROTINAARQUIVO = R.Handle)")
  	SQL.Add(" WHERE R.TABTIPO  = 1")
  	SQL.Add("   And R.SITUACAO In ('4', '5', '6', '7', '8' )")
  	SQL.Add("   AND D.HANDLE   = :HDOCUMENTO")
  	SQL.ParamByName("HDOCUMENTO").AsInteger = qDocumento.FieldByName("HANDLE").AsInteger
	SQL.Active = True

	If SQL.FieldByName("QTDE").AsInteger > 0 Then
    	Err.Raise(vbsUserException,"","Não é possível alterar o documento. Existe uma rotina de remessa processada!")
    	CanContinue = False
    	Set qDocumento = Nothing
		Set SQL = Nothing
    	Exit Sub
  	End If

  	Set qDocumento = Nothing
	Set SQL = Nothing

End Sub
