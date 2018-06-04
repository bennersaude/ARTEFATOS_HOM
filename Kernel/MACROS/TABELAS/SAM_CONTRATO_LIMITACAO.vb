'HASH: 9EB0E861E65E3C5F133354CC875F6408
'Macro: SAM_CONTRATO_LIMITACAO
'#Uses "*bsShowMessage"
'#Uses "*TipoPeriodoLimiteValido"

Dim gPlano			As Integer
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
		Sql.Add("  FROM SAM_PLANO_LIMITACAO                              ")
		Sql.Add(" WHERE LIMITACAO = :pLIMITACAO                          ")
		Sql.Add("   AND PLANO = (SELECT PLANO                            ")
		Sql.Add("                  FROM SAM_CONTRATO                     ")
		Sql.Add("                 WHERE HANDLE = :pCONTRATO)             ")
		Sql.ParamByName("pLIMITACAO").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
		Sql.ParamByName("pCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
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
  If CurrentQuery.State = 3 Then
    Sql.Active = False
    Sql.Clear
    Sql.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
    Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
    Sql.Active = True

    If Sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
      PERIODO.Visible = False
    Else
      PERIODO.Visible = True
    End If

    Set Sql = Nothing
  End If
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

  If CurrentQuery.State = 1 Then
    TIPOLIMITACAO.ReadOnly = True
  Else
    TIPOLIMITACAO.ReadOnly = False
  End If

  bIntercambiavel = CurrentQuery.FieldByName("INTERCAMBIAVEL").AsString = "S"

  'SMS 61198 - Matheus - Início
  If CurrentQuery.State = 3 Then
    Dim Sql As Object
    Set Sql = NewQuery

    Sql.Active = False
    Sql.Clear
    Sql.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
    Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
    Sql.Active = True

    If Sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
      PERIODO.Visible = False
    Else
      PERIODO.Visible = True
    End If

    Set Sql = Nothing
  End If
  'SMS 61198 - Matheus - Fim

  'SMS 95433 - Paulo Melo - 11/04/2008
  If CurrentQuery.FieldByName("LIMITACAO").AsInteger > 0 Then
    Dim Sql2 As Object
    Set Sql2 = NewQuery
    Sql2.Active = False
    Sql2.Clear
    Sql2.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
    Sql2.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
    Sql2.Active = True

    If Sql2.FieldByName("PERIODICIDADE").AsInteger = 2 Then  ' se for em semanas
      PERIODO.Visible = False                                ' esconde o campo Periodo (que só vale pra meses)
    Else
      PERIODO.Visible = True
    End If

    Set Sql2 = Nothing
  End If
  'SMS 95433 - Paulo Melo - 11/04/2008 - FIM

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim vHandle As Integer

  If CurrentQuery.State = 1 Then
    TIPOLIMITACAO.ReadOnly = True
  Else
    TIPOLIMITACAO.ReadOnly = False
  End If

  If (VisibleMode Or WebMode) Then
    vHandle = RecordHandleOfTable("SAM_CONTRATO")
  End If

  If WebMode Then
  	PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = " + Str(vHandle) + ")"
  ElseIf VisibleMode Then
  	PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = " + Str(vHandle) + ")"
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim vHandle As Integer

  If (VisibleMode Or WebMode) Then
    vHandle = RecordHandleOfTable("SAM_CONTRATO")
  End If

  If WebMode Then
  	PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = " + Str(vHandle) + ")"
  ElseIf VisibleMode Then
  	PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = " + Str(vHandle) + ")"
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String
  Dim Sql As Object

  Condicao = "AND LIMITACAO = " + CurrentQuery.FieldByName("LIMITACAO").AsString
  Condicao = Condicao + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString 'Anderson sms 21638
  Condicao = Condicao + " AND TIPOCONTAGEM = '" + CurrentQuery.FieldByName("TIPOCONTAGEM").AsString + "'"

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_LIMITACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha <> "" Then
    CanContinue = False
    bsShowMessage(Linha, "E")
    Exit Sub
  Else
    CanContinue = True
  End If

  'Caso tenha outra limitação com vigência em aberto,permitir continuar caso não tenha limite sem módulo.
  Set Sql = NewQuery
  Sql.Add("Select A.LIMITACAO, ")
  Sql.Add("       B.CONTRATOMODULO ")
  Sql.Add("  FROM SAM_CONTRATO_LIMITACAO A ")
  Sql.Add("  Left Join SAM_CONTRATO_LIMITACAO_MOD B On B.CONTRATOLIMITACAO = A.HANDLE ")
  Sql.Add(" WHERE A.CONTRATO = :CONTRATO ")
  Sql.Add("   AND A.LIMITACAO = :LIMITACAO ")
  Sql.Add("   AND A.PLANO = :PLANO ")'Anderson sms 21638
  Sql.Add("   AND B.CONTRATOMODULO IS NULL ")
  Sql.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  Sql.ParamByName("LIMITACAO").Value = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  Sql.ParamByName("PLANO").Value = CurrentQuery.FieldByName("PLANO").AsInteger 'Anderson sms 21638
  Sql.Active = True
  'End If
  If Sql.EOF Then
    If(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >CurrentQuery.FieldByName("DATAFINAL").AsDateTime)And _
       (CurrentQuery.FieldByName("DATAFINAL").AsDateTime >0)Then
    CanContinue = False
    bsShowMessage("A data final não pode ser anterior a data inicial!", "E")
    Exit Sub
  Else
    CanContinue = True
  End If
  End If


  CanContinue = TipoPeriodoLimiteValido(CurrentQuery.FieldByName("TIPOCONTAGEM").AsString, CurrentQuery.FieldByName("TIPOPERIODO").AsString)

	If (bIntercambiavel And CurrentQuery.FieldByName("TIPOCONTAGEM").AsString <> "C") Then
		bsShowMessage("Para limitações intercambiáveis, a contagem deve ser feita no 'Contrato'.", "E")
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
  If CurrentQuery.State = 3 Then
    Sql.Active = False
    Sql.Clear
    Sql.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
    Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
    Sql.Active = True

    If Sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then  CurrentQuery.FieldByName("PERIODO").AsInteger = 1

    Set Sql = Nothing
  End If
  'SMS 61198 - Matheus - Fim
End Sub


Public Function CheckVigencia As Boolean
  CheckVigencia = True
  Dim Sql As Object
  Set Sql = NewQuery
  Sql.Add("SELECT DATAADESAO, DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  Sql.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  Sql.Active = True
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <Sql.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data Inicial de Limitação menor que a Adesão do Contrato!", "E")
    CheckVigencia = False
  Else
    If Not Sql.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >Sql.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data Inicial de Limitação maior que o cancelamento do Contrato!", "E")
        CheckVigencia = False
      End If
    End If
  End If
  Set Sql = Nothing
End Function

