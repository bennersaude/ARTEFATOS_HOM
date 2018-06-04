'HASH: 141CD3170C480F5A63B3BDA1293B67C4
'Macro: SAM_BENEFICIARIO_LIMITACAO
'#Uses "*bsShowMessage"

Public Sub LIMITACAO_OnChange()
	Dim Nulo As String
	Set Sql = NewQuery

	Nulo = "ISNULL"
	If (StrPos("ORACLE", SQLServer) > 0) Then
		Nulo = "NVL"
	Else
		If (StrPos("CACHE", SQLServer) > 0) Then
			Nulo = "NVL"
		Else
			If (StrPos("DB2", SQLServer) > 0) Then
				Nulo = "COALESCE"
			End If
		End If
	End If

	If Not CurrentQuery.FieldByName("LIMITACAO").IsNull Then
	    SQL.Active = False
	    SQL.Clear
		SQL.Add("SELECT " + Nulo + "(PERIODO, 0) PERIODO,                ")
		SQL.Add("       " + Nulo + "(TIPOLIMITACAO, 'A') TIPOLIMITACAO,  ")
		SQL.Add("       " + Nulo + "(INTERCAMBIAVEL, 'N') INTERCAMBIAVEL,")
		SQL.Add("       " + Nulo + "(TIPOCONTAGEM, 'B') TIPOCONTAGEM,    ")
		SQL.Add("       " + Nulo + "(TIPOPERIODO, 'C') TIPOPERIODO,      ")
		SQL.Add("       " + Nulo + "(TABTIPOLIMITE, 1) TABTIPOLIMITE,    ")
		SQL.Add("       QTDLIMITE,                                       ")
		SQL.Add("       " + Nulo + "(TABTIPOVALOR, 1) TABTIPOVALOR,      ")
		SQL.Add("       TABELAUS,                                        ")
		SQL.Add("       VLRLIMITE                                        ")
		SQL.Add("  FROM SAM_CONTRATO_LIMITACAO                       ")
		SQL.Add(" WHERE HANDLE = :pLIMITACAO                         ")
		SQL.ParamByName("pLIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
		SQL.Active = True

		If (Not SQL.EOF) Then
			CurrentQuery.FieldByName("PERIODOCONTAGEM").Value	= SQL.FieldByName("PERIODO").Value
			CurrentQuery.FieldByName("TIPOLIMITACAO").Value		= SQL.FieldByName("TIPOLIMITACAO").Value
			CurrentQuery.FieldByName("TIPOPERIODO").Value		= SQL.FieldByName("TIPOPERIODO").Value
			CurrentQuery.FieldByName("TABTIPOLIMITE").Value		= SQL.FieldByName("TABTIPOLIMITE").Value
			CurrentQuery.FieldByName("QTDLIMITE").Value			= SQL.FieldByName("QTDLIMITE").Value
			CurrentQuery.FieldByName("TABTIPOVALOR").Value		= SQL.FieldByName("TABTIPOVALOR").Value
			CurrentQuery.FieldByName("TABELAUS").Value			= SQL.FieldByName("TABELAUS").Value
			CurrentQuery.FieldByName("VLRLIMITE").Value			= SQL.FieldByName("VLRLIMITE").Value
		End If
	End If

  'SMS 61198 - Matheus - Início
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SL.PERIODICIDADE          ")
  SQL.Add("  FROM SAM_LIMITACAO SL,         ")
  SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
  SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODOCONTAGEM.Visible = False
  Else
    PERIODOCONTAGEM.Visible = True
  End If

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
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
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SL.PERIODICIDADE          ")
  SQL.Add("  FROM SAM_LIMITACAO SL,         ")
  SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
  SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODOCONTAGEM.Visible = False
  Else
    PERIODOCONTAGEM.Visible = True
  End If

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_AfterInsert()

  Dim qBeneficiario           As Object
  Dim qDataPadrao             As Object
  Dim vsNvl                   As String
  Dim vsCondicao              As String
  Dim vDataInfinita			  As String
  Dim vDataPadrao			  As String

  Set qBeneficiario = NewQuery

  Set qDataPadrao = NewQuery
  qDataPadrao.Clear
  qDataPadrao.Add("SELECT VALOR                          ")
  qDataPadrao.Add("  FROM Z_VARIAVEIS                    ")
  qDataPadrao.Add(" WHERE NOME LIKE 'DATABASEDATEFORMAT' ")
  qDataPadrao.Active = True

  If qDataPadrao.FieldByName("VALOR").AsString <> "" Then
    vDataPadrao = qDataPadrao.FieldByName("VALOR").AsString
  Else
    vDataPadrao = "DD/MM/YYYY"
  End If

  Set qDataPadrao = Nothing

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

  If (InStr(SQLServer, "MSSQL") > 0) Then
    vsNvl             = "ISNULL"
  ElseIf (InStr(SQLServer, "ORACLE") > 0) Then
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

  If WebMode Then
  	vsCondicao = vsCondicao + " AND DATAINICIAL <= @CAMPO(DATAINICIAL)"
  	vsCondicao = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= @CAMPO(DATAINICIAL)"
  ElseIf VisibleMode Then

	If (InStr(SQLServer, "MSSQL") > 0) Then
  		vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  	Else
  		vsCondicao = vsCondicao + " AND DATAINICIAL <= TO_DATE(@DATAINICIAL , '" + vDataPadrao + "' )"
  	End If

  	vsCondicao = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= " + SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  End If

  'As verificações feitas acima também serão feitas antes de gravar o registro, pois o usuário pode selecionar uma tabela PF evento e depois alterar as datas da vigência.
  If WebMode Then
  	  LIMITACAO.WebLocalWhere = vsCondicao
  ElseIf VisibleMode Then
	  LIMITACAO.LocalWhere = vsCondicao
  End If
End Sub



Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = "AND LIMITACAO = " + CurrentQuery.FieldByName("LIMITACAO").AsString
  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")
  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_LIMITACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)
  If Linha <> "" Then
	CanContinue = False
	bsShowMessage(Linha, "E")
	Exit Sub
  End If
  CanContinue = CheckVigenciaBenef

  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT SL.PERIODICIDADE          ")
  SQL.Add("  FROM SAM_LIMITACAO SL,         ")
  SQL.Add("       SAM_CONTRATO_LIMITACAO SCL")
  SQL.Add(" WHERE SCL.HANDLE = :HANDLE      ")
  SQL.Add("   AND SL.HANDLE = SCL.LIMITACAO ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then  CurrentQuery.FieldByName("PERIODOCONTAGEM").AsInteger = 1

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Function CheckVigenciaBenef As Boolean
  CheckVigenciaBenef = True
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DATAADESAO,DATACANCELAMENTO FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEF")
  SQL.ParamByName("BENEF").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  SQL.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial da Limitação inferior a Adesão do Beneficiário!", "E")
    CheckVigenciaBenef = False
  Else
    If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >SQL.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data de Inicial da Limitação maior que o cancelamento do Beneficiário !", "E")
        CheckVigenciaBenef = False
      End If
    End If
  End If
  Set SQL = Nothing
End Function

