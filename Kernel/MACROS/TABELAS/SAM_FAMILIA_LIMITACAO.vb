'HASH: EE57B1E1470113B6E791783C8C3CF079
'Macro: SAM_FAMILIA_LIMITACAO
'#Uses "*bsShowMessage"
'#Uses "*TipoPeriodoLimiteValido"

Dim bIntercambiavel	As Boolean

Public Sub INTERCAMBIAVEL_OnChange()
	bIntercambiavel = Not bIntercambiavel
End Sub

Public Sub LIMITACAO_OnChange()
	Dim Sql As Object
	Dim Nulo As String
	Set Sql = NewQuery

	Nulo = "ISNULL"
	If (StrPos("ORACLE", SQLServer) > 0) Or (StrPos("CACHE", SQLServer) > 0) Then
		Nulo = "NVL"
	Else
		If (StrPos("DB2", SQLServer) > 0) Then
			Nulo = "COALESCE"
		End If
	End If

	If Not CurrentQuery.FieldByName("LIMITACAO").IsNull Then
	    Sql.Active = False
	    Sql.Clear
		Sql.Add("SELECT " + Nulo + "(PERIODO, 0) PERIODO,                ")
		Sql.Add("       " + Nulo + "(TIPOLIMITACAO, 'A') TIPOLIMITACAO,  ")
		Sql.Add("       " + Nulo + "(INTERCAMBIAVEL, 'N') INTERCAMBIAVEL,")
		Sql.Add("       " + Nulo + "(TIPOCONTAGEM, 'B') TIPOCONTAGEM,    ")
		Sql.Add("       " + Nulo + "(TIPOPERIODO, 'C') TIPOPERIODO,      ")
		Sql.Add("       " + Nulo + "(TABTIPOLIMITE, 1) TABTIPOLIMITE,    ")
		Sql.Add("       QTDLIMITE,                                       ")
		Sql.Add("       " + Nulo + "(TABTIPOVALOR, 1) TABTIPOVALOR,      ")
		Sql.Add("       TABELAUS,                                        ")
		Sql.Add("       VLRLIMITE                                        ")
		Sql.Add("  FROM SAM_CONTRATO_LIMITACAO                       ")
		Sql.Add(" WHERE HANDLE = :pLIMITACAO                         ")
		Sql.ParamByName("pLIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
		Sql.Active = True

		If (Not Sql.EOF) Then
			CurrentQuery.FieldByName("PERIODO").Value			= Sql.FieldByName("PERIODO").Value
			CurrentQuery.FieldByName("TIPOLIMITACAO").Value		= Sql.FieldByName("TIPOLIMITACAO").Value
			CurrentQuery.FieldByName("INTERCAMBIAVEL").Value	= Sql.FieldByName("INTERCAMBIAVEL").Value
			CurrentQuery.FieldByName("TIPOCONTAGEM").Value		= Sql.FieldByName("TIPOCONTAGEM").Value
			CurrentQuery.FieldByName("TIPOPERIODO").Value		= Sql.FieldByName("TIPOPERIODO").Value
			CurrentQuery.FieldByName("TABTIPOLIMITE").Value		= Sql.FieldByName("TABTIPOLIMITE").Value
			CurrentQuery.FieldByName("QTDLIMITE").Value			= Sql.FieldByName("QTDLIMITE").Value
			CurrentQuery.FieldByName("TABTIPOVALOR").Value		= Sql.FieldByName("TABTIPOVALOR").Value
			CurrentQuery.FieldByName("TABELAUS").Value			= Sql.FieldByName("TABELAUS").Value
			CurrentQuery.FieldByName("VLRLIMITE").Value			= Sql.FieldByName("VLRLIMITE").Value
		End If
	End If

  'SMS 61198 - Matheus - Início
  Sql.Active = False
  Sql.Clear
  Sql.Add("SELECT SL.PERIODICIDADE          ")
  Sql.Add("  FROM SAM_LIMITACAO SL,         ")
  Sql.Add("       SAM_CONTRATO_LIMITACAO SCL")
  Sql.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  Sql.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  Sql.Active = True

  If Sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODO.Visible = False
  Else
    PERIODO.Visible = True
  End If

  Set Sql = Nothing
  'SMS 61198 - Matheus - Fim

End Sub



Public Sub TABLE_AfterEdit()
  Dim qBeneficiario           As Object
  Dim vsNvl                   As String
  Dim vsCondicao              As String
  Dim vDataInfinita			  As String


  Set qBeneficiario = NewQuery

  'Selecionar as datas de adesão e cancelamento do beneficiário e formatá-las de acordo com o banco de dados.
  qBeneficiario.Clear
  qBeneficiario.Add("SELECT DATAADESAO,     ")
  qBeneficiario.Add("       DATACANCELAMENTO")
  qBeneficiario.Add("  FROM SAM_BENEFICIARIO")
  qBeneficiario.Add(" WHERE HANDLE = :HANDLE")
  qBeneficiario.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO")
  qBeneficiario.Active = True

  vsDataCancelamentoBenef = ""
  'SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)             = ""


  If (InStr(SQLServer, "MSSQL") > 0) Then
    vsNvl             = "ISNULL"
  ElseIf (InStr(SQLServer, "ORACLE") > 0) Then
    vsNvl             = "NVL"
  ElseIf (InStr(SQLServer, "DB2") > 0) Then
    vsNvl             = "COALESCE"
  ElseIf (InStr(SQLServer, "CACHE") > 0) Then
    vsNvl             = "NVL"
  End If

  vDataInfinita = SQLAddYear(SQLDate(ServerDate),"200")

  'Serão exibidos apenas os registros onde a data final é maior ou igual à adesão do beneficiário ou onde a vigência ainda está aberta.
  vsCondicao = vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= " + SQLDate(qBeneficiario.FieldByName("DATAADESAO").AsDateTime)
  If (Not qBeneficiario.FieldByName("DATACANCELAMENTO").IsNull) Then
    'Se o beneficiário possui data de cancelamento, o filtro ficará restrito àqueles registros onde a data inicial é menor ou igual ao cancelamento do beneficiário.
    vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(qBeneficiario.FieldByName("DATACANCELAMENTO").AsDateTime)
  End If
  Set qBeneficiario = Nothing

  'Filtrar também pela vigência do próprio registro que está sendo incluído/alterado.



  'As verificações feitas acima também serão feitas antes de gravar o registro, pois o usuário pode selecionar uma tabela PF evento e depois alterar as datas da vigência.
   If WebMode Then
    vsCondicao = vsCondicao + " AND DATAINICIAL <= @CAMPO(DATAINICIAL)"
  	LIMITACAO.WebLocalWhere = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= @CAMPO(DATAINICIAL)"
  ElseIf VisibleMode Then
    vsCondicao = vsCondicao + " AND DATAINICIAL <= @DATAINICIAL"
    LIMITACAO.LocalWhere = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= @DATAINICIAL"
  End If

End Sub

Public Sub TABLE_AfterInsert()
  Dim qFamilia                As Object
  Dim vsNvl                   As String
  Dim vsCondicao              As String
  Dim vDataInfinita			  As String


  Set qFamilia = NewQuery

  'Selecionar as datas de adesão e cancelamento do beneficiário e formatá-las de acordo com o banco de dados.
  qFamilia.Clear
  qFamilia.Add("SELECT DATAADESAO,     ")
  qFamilia.Add("       DATACANCELAMENTO")
  qFamilia.Add("  FROM SAM_FAMILIA")
  qFamilia.Add(" WHERE HANDLE = :HANDLE")
  qFamilia.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  qFamilia.Active = True

  vsDataCancelamentoBenef = ""
  'SQLDate(CurrentQuery.FieldByName("DATAFINAL").AsDateTime)             = ""


  If (InStr(SQLServer, "MSSQL") > 0) Then
    vsNvl             = "ISNULL"
  ElseIf (InStr(SQLServer, "ORACLE") > 0) Then
    vsNvl             = "NVL"
  ElseIf (InStr(SQLServer, "DB2") > 0) Then
    vsNvl             = "COALESCE"
  ElseIf (InStr(SQLServer, "CACHE") > 0) Then
    vsNvl             = "NVL"
  End If

  vDataInfinita = SQLAddYear(SQLDate(ServerDate),"200")

  'Serão exibidos apenas os registros onde a data final é maior ou igual à adesão do beneficiário ou onde a vigência ainda está aberta.
  If WebMode Then
  	vsCondicao = vsNvl + "(A.DATAFINAL, " + vDataInfinita + ") >= " + SQLDate(qFamilia.FieldByName("DATAADESAO").AsDateTime)
  	If (Not qFamilia.FieldByName("DATACANCELAMENTO").IsNull) Then
	    vsCondicao = vsCondicao + " AND A.DATAINICIAL <= " + SQLDate(qFamilia.FieldByName("DATACANCELAMENTO").AsDateTime)
  	End If
  	Set qFamilia = Nothing

  ElseIf VisibleMode Then
  	vsCondicao = vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= " + SQLDate(qFamilia.FieldByName("DATAADESAO").AsDateTime)
  	If (Not qFamilia.FieldByName("DATACANCELAMENTO").IsNull) Then
	    vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(qFamilia.FieldByName("DATACANCELAMENTO").AsDateTime)
  	End If
  	Set qFamilia = Nothing

  End If

  'Filtrar também pela vigência do próprio registro que está sendo incluído/alterado.



  'As verificações feitas acima também serão feitas antes de gravar o registro, pois o usuário pode selecionar uma tabela PF evento e depois alterar as datas da vigência.
   If WebMode Then
    vsCondicao = vsCondicao + " AND A.DATAINICIAL <= @CAMPO(DATAINICIAL)"
  	'LIMITACAO.WebLocalWhere = vsCondicao + " AND " + vsNvl + "(A.DATAFINAL, " + vDataInfinita + ") >= @CAMPO(DATAINICIAL)"
  	LIMITACAO.WebLocalWhere = "A.DATAINICIAL <= @CAMPO(DATAINICIAL) " ' AND ISNULL(A.DATAFINAL, '20200101') >= @CAMPO(DATAINICIAL)"
  ElseIf VisibleMode Then
    vsCondicao = vsCondicao + " AND DATAINICIAL <= @DATAINICIAL"
    LIMITACAO.LocalWhere = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= @DATAINICIAL"
  End If

End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If

  'SMS 61198 - Matheus - Início
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Active = False
  Sql.Clear
  Sql.Add("SELECT SL.PERIODICIDADE          ")
  Sql.Add("  FROM SAM_LIMITACAO SL,         ")
  Sql.Add("       SAM_CONTRATO_LIMITACAO SCL")
  Sql.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  Sql.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  Sql.Active = True

  If Sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODO.Visible = False
  Else
    PERIODO.Visible = True
  End If

  Set Sql = Nothing
  'SMS 61198 - Matheus - Fim

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Condicao As String

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
	Condicao = " AND LIMITACAO = " + CurrentQuery.FieldByName("LIMITACAO").AsString
	Condicao = Condicao + " AND TIPOCONTAGEM = '" + CurrentQuery.FieldByName("TIPOCONTAGEM").AsString + "'"
	Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_LIMITACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FAMILIA", Condicao)
	Set Interface = Nothing
	If Linha <> "" Then
		CanContinue = False
		bsShowMessage(Linha, "E")
		Exit Sub
	End If

	CanContinue = TipoPeriodoLimiteValido(CurrentQuery.FieldByName("TIPOCONTAGEM").AsString, CurrentQuery.FieldByName("TIPOPERIODO").AsString)
	If Not CanContinue Then Exit Sub

	If (bIntercambiavel And CurrentQuery.FieldByName("TIPOCONTAGEM").AsString <> "F") Then
		bsShowMessage("Para limitações intercambiáveis, a contagem deve ser feita na 'Família'.", "E")
		CanContinue = False
		Exit Sub
	End If

	CanContinue = CheckVigencia

    'Luciano T. Alberti - SMS 53386 - 06/03/2006 - Início
    If (CurrentQuery.FieldByName("SOMARSALARIO").AsString = "S") And _
       (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString  <> "B") Then
      bsShowMessage("Somar salário somente para contagem no beneficiário", "E")
      CanContinue = False
    End If
    'Luciano T. Alberti - SMS 53386 - 06/03/2006 - Fim


  'SMS 61198 - Matheus - Início
  Dim Sql As Object
  Set Sql = NewQuery

  Sql.Active = False
  Sql.Clear
  Sql.Add("SELECT SL.PERIODICIDADE,         ")
  Sql.Add("       SCL.DATAFINAL          ")
  Sql.Add("  FROM SAM_LIMITACAO SL,         ")
  Sql.Add("       SAM_CONTRATO_LIMITACAO SCL")
  Sql.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  Sql.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  Sql.Active = True

  If Sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    CurrentQuery.FieldByName("PERIODO").AsInteger = 1
  End If

  If ((Not Sql.FieldByName("DATAFINAL").IsNull) And (Sql.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAFINAL").AsDateTime)) Or _
      ((CurrentQuery.FieldByName("DATAFINAL").IsNull) And (Not Sql.FieldByName("DATAFINAL").IsNull)) Then
  	bsShowMessage("Data Final de vigência deve ser menor ou igual a data final da limitação do contrato. Data Final: " + FormatDateTime2("dd/mm/yyyy", CurrentQuery.FieldByName("DATAFINAL").AsDateTime), "E")
    CanContinue = False
    Exit Sub
  End If

  Set Sql = Nothing

End Sub


Public Function CheckVigencia As Boolean
  CheckVigencia = True
  Dim Sql As Object
  Set Sql = NewQuery
  Sql.Add("SELECT * FROM SAM_FAMILIA WHERE HANDLE = :FAMILIA")
  Sql.ParamByName("FAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
  Sql.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <Sql.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial de Limitação menor que a Adesão da Família!", "E")
    CheckVigencia = False
  Else
    If Not Sql.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >Sql.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data Inicial de Limitação maior que o cancelamento da Família!", "E")
        CheckVigencia = False
      End If
    End If
  End If
  Set Sql = Nothing
End Function

