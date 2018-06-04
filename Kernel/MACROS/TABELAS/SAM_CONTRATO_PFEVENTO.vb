'HASH: F16F2A72747650727B4455A6C3A7FD08
'Macro: SAM_CONTRATO_PFEVENTO
'#Uses "*bsShowMessage"
' Mauricio Ibelli -sms 2075 -27/03/2001 -Inibido condicao para considerar qdo e tipo de pf
Dim gPlano As Long

Public Sub PLANO_OnChange()
  'Anderson sms 21638
  gPlano = CurrentQuery.FieldByName("PLANO").AsInteger
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll

  Dim qUpdateContrato As BPesquisa
  Set qUpdateContrato = NewQuery

  If (CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 4) Then
    qUpdateContrato.Clear
    qUpdateContrato.Add("UPDATE SAM_CONTRATO SET USAPFPORSALARIO = 'S' WHERE HANDLE = :HANDLE")
    qUpdateContrato.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
    qUpdateContrato.ExecSQL
  End If
  Set qUpdateContrato = Nothing

End Sub


Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
  End If
  'Anderson sms 21638
  gPlano = CurrentQuery.FieldByName("PLANO").AsInteger

  If WebMode Then
    If CurrentQuery.FieldByName("REGRA").AsInteger = 0 Then
      PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CAMPO(CONTRATO))"
    Else
      PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = (SELECT CONTRATO FROM SAM_CONTRATO_PFREGRA WHERE HANDLE = @CAMPO(REGRA)))"
    End If
  ElseIf VisibleMode Then
    PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_PLANO WHERE CONTRATO = @CONTRATO)"
  End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qContrato As Object
  Set qContrato = NewQuery

  qContrato.Clear
  qContrato.Add("SELECT DATAADESAO,     ")
  qContrato.Add("       DATACANCELAMENTO")
  qContrato.Add("  FROM SAM_CONTRATO    ")
  qContrato.Add(" WHERE HANDLE = :HANDLE")
  If (VisibleMode) Or (WebMode) Then
    qContrato.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO")
  Else
    qContrato.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  End If
  qContrato.Active = True

  'Não permitir a inclusão de registros com vigência iniciando antes da adesão do contrato ou terminando depois do cancelamento do mesmo (se houver).
  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < qContrato.FieldByName("DATAADESAO").AsDateTime) Then
    bsShowMessage("A data inicial não pode ser anterior à adesão do contrato.", "E")
    CanContinue = False
    Set qContrato = Nothing
    Exit Sub
  Else
    If (Not qContrato.FieldByName("DATACANCELAMENTO").IsNull) Then
      If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
        bsShowMessage("A data final não pode ficar em aberto, pois o contrato possui data de cancelamento.", "E")
        CanContinue = False
        Set qContrato = Nothing
        Exit Sub
      Else
        If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime > qContrato.FieldByName("DATACANCELAMENTO").AsDateTime) Then
          bsShowMessage("A data final não pode ser posterior ao cancelamento do contrato.", "E")
          CanContinue = False
          Set qContrato = Nothing
          Exit Sub
        End If
      End If
    End If
  End If
  Set qContrato = Nothing

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Dim vHandleContratoPFREGRA As Integer
  vHandleContratoPFREGRA = 0
  If (VisibleMode) Or (WebMode) Then
    vHandleContratoPFREGRA = RecordHandleOfTable("SAM_CONTRATO_PFREGRA")
  Else
    vHandleContratoPFREGRA = CurrentQuery.FieldByName("REGRA").AsInteger
  End If

  If vHandleContratoPFREGRA > 0 Then
    Condicao = "AND TABELAPFEVENTO = " + CurrentQuery.FieldByName("TABELAPFEVENTO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString + " AND REGRA = " + Str(vHandleContratoPFREGRA)
  Else
    Condicao = "AND TABELAPFEVENTO = " + CurrentQuery.FieldByName("TABELAPFEVENTO").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString + " AND REGRA IS NULL" 'Anderson sms 21638
  End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PFEVENTO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

  'SMS 40817 - Anderson Lonardoni - 03/05/2005
  'SE O TIPO DE CONTAGEM FOR NO CONTRATO O TIPO DE PERÍODO DEVERÁ SER CIVIL OU POR ADESÃO DO CONTRATO
  If (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "C") _
       And ((CurrentQuery.FieldByName("TIPOPERIODO").AsString = "F") Or (CurrentQuery.FieldByName("TIPOPERIODO").AsString = "B")) _
       And (CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 2) Then
    CanContinue = False
    bsShowMessage("Tipo de contagem no contrato exige que o tipo de período seja civil ou por adesão do contrato!", "E")
    Exit Sub
  End If
  'SE O TIPO DE CONTAGEM FOR NA FAMÍLIA O TIPO DE PERÍODO DEVERÁ SER CIVIL, POR ADESÃO DO CONTRATO OU DA FAMÍLIA
  If (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "F") And (CurrentQuery.FieldByName("TIPOPERIODO").AsString = "B") Then
    CanContinue = False
    bsShowMessage("Tipo de contagem na família exige que o tipo de período seja civil, por adesão do contrato ou da família!", "E")
    Exit Sub
  End If

  'Andreia sms 23013
  Set qContrato = NewQuery
  qContrato.Active = False
  qContrato.Clear
  qContrato.Add("SELECT C.PERMITEPFNAIMPORTACAOBENEF ")
  qContrato.Add(" FROM SAM_CONTRATO C")
  qContrato.Add(" WHERE C.HANDLE = :HCONTRATO")
  qContrato.ParamByName("HCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qContrato.Active = True

  If ((CurrentQuery.FieldByName("ATUALIZARNAIMPORTACAO").AsString = "S") And (qContrato.FieldByName("PERMITEPFNAIMPORTACAOBENEF").AsString = "N")) Then
    bsShowMessage("Contrato não permite PF na importação de beneficiários. Parâmetro 'Atualizar na importação' não deve estar marcado", "E")
    Set qContrato = Nothing
    CanContinue = False
    Exit Sub
  End If
  Set qContrato = Nothing
  'sms 23013

  If ((CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 2) And _
       (CurrentQuery.FieldByName("TABPADRAOCONTAGEM").AsInteger) And _
       (CurrentQuery.FieldByName("TIPOPFVARIAVEL").AsString = "V") And _
       (CurrentQuery.FieldByName("TABELAUS").IsNull)) Then
    CanContinue = False
    TABELAUS.SetFocus
    bsShowMessage("Tabela de US obrigatória para 'Tipo da PF variável' por valor.", "E")
    Exit Sub
  End If

  If ((Not CurrentQuery.FieldByName("DATAFINAL").IsNull) And _
       (DateAdd("m", CurrentQuery.FieldByName("PERIODO").AsInteger, _
       CurrentQuery.FieldByName("DATAINICIAL").AsDateTime) > _
       CurrentQuery.FieldByName("DATAFINAL").AsDateTime) And _
       CurrentQuery.FieldByName("TIPOPERIODO").AsString = "P" And _
       CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 2 And _
       CurrentQuery.FieldByName("TABPADRAOCONTAGEM").AsInteger = 1) Then
    CanContinue = False
    bsShowMessage("Para a configuração do tipo de período ""Vigência da PF""," + Chr(13) + _
           "a Data inicial somada ao núm. de meses do Período" + Chr(13) + _
           "deve ser inferior ou igual à Data final." + Chr(13) + Chr(13) + _
           "Data inicial : " + CurrentQuery.FieldByName("DATAINICIAL").AsString + Chr(13) + _
           "Período : " + CurrentQuery.FieldByName("PERIODO").AsString + " mes(es)" + Chr(13) + _
           "Data final : " + CurrentQuery.FieldByName("DATAFINAL").AsString + Chr(13) + _
           "Data final resultante : " + Format(DateAdd("m", CurrentQuery.FieldByName("PERIODO").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)), "E")
    DATAFINAL.SetFocus
    Exit Sub
  End If


  If (CurrentQuery.FieldByName("INTERCAMBIAVEL").AsString = "S" And CurrentQuery.FieldByName("TIPOCONTAGEM").AsString <> "C") Then
	bsShowMessage("Para PFs intercambiáveis, a contagem deve ser feita no 'Contrato'.", "E")
	CanContinue = False
	Exit Sub
  End If


  If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <= ServerDate) Then
      bsShowMessage("Não é possível fechar vigência retroativa, favor fechar com a data de amanhã", "E")
      CanContinue = False
      Exit Sub
    End If
  End If


  Dim callEntity As CSEntityCall
  Dim retorno As String

  If ((CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 5) And (CurrentQuery.FieldByName("VALORDEMONSFIXOPF").AsInteger = 0 And CurrentQuery.FieldByName("VALORDEMONSPERCENTUAL").AsInteger = 0))Then
	bsShowMessage("Para Pf Demonstrada informar ou valor fixo ou valor percentual.", "E")
	VALORDEMONSFIXOPF.SetFocus
  	CanContinue = False
  	Exit Sub
  End If

  If (CurrentQuery.InInsertion And (CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 5))Then
  	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.Beneficiarios.Contrato.SamContratoPfEvento, Benner.Saude.Entidades", "Validar")
  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("CONTRATO").AsInteger)
  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger)
  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("PLANO").AsInteger)
  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  	callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)
  	retorno = callEntity.Execute
  End If

  If (retorno <> "") Then
	bsShowMessage("A participação financeira em questão possui evento em duplicidade. Verificar os grupos de participações financeira: " + retorno + ".", "E")
	TABELAPFEVENTO.SetFocus
	CanContinue = False
	Set callEntity = Nothing
	Exit Sub
  End If

  Set callEntity = Nothing

End Sub

Public Sub TABELAPFEVENTO_OnChange()
  Dim qBuscaDadosPlano As Object
  Set qBuscaDadosPlano = NewQuery

  qBuscaDadosPlano.Clear
  qBuscaDadosPlano.Add("SELECT * FROM SAM_PLANO_PFEVENTO WHERE PLANO = :pPLANO AND TABELAPFEVENTO = :pPFEVENTO")
  qBuscaDadosPlano.ParamByName("pPLANO").AsInteger = gPlano
  qBuscaDadosPlano.ParamByName("pPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
  qBuscaDadosPlano.Active = True

  If (Not qBuscaDadosPlano.EOF) Then
    If (qBuscaDadosPlano.FieldByName("ACEITAFINANCIAMENTO").IsNull) Then
      CurrentQuery.FieldByName("ACEITAFINANCIAMENTO").AsString = "N"
    Else
      CurrentQuery.FieldByName("ACEITAFINANCIAMENTO").Value = qBuscaDadosPlano.FieldByName("ACEITAFINANCIAMENTO").Value
    End If

    If (qBuscaDadosPlano.FieldByName("ACEITAPARCELAMENTO").IsNull) Then
      CurrentQuery.FieldByName("ACEITAPARCELAMENTO").AsString = "N"
    Else
      CurrentQuery.FieldByName("ACEITAPARCELAMENTO").Value = qBuscaDadosPlano.FieldByName("ACEITAPARCELAMENTO").Value
    End If

    If (qBuscaDadosPlano.FieldByName("INTERCAMBIAVEL").IsNull) Then
      CurrentQuery.FieldByName("INTERCAMBIAVEL").AsString = "N"
    Else
      CurrentQuery.FieldByName("INTERCAMBIAVEL").Value = qBuscaDadosPlano.FieldByName("INTERCAMBIAVEL").Value
    End If

    If (qBuscaDadosPlano.FieldByName("TABPADRAOCONTAGEM").IsNull) Then
      CurrentQuery.FieldByName("TABPADRAOCONTAGEM").AsInteger = 1
    Else
      CurrentQuery.FieldByName("TABPADRAOCONTAGEM").Value = qBuscaDadosPlano.FieldByName("TABPADRAOCONTAGEM").Value
    End If

    CurrentQuery.FieldByName("CODIGOPF").Value = qBuscaDadosPlano.FieldByName("CODIGOPF").Value
    CurrentQuery.FieldByName("PERIODO").Value = qBuscaDadosPlano.FieldByName("PERIODO").Value
    CurrentQuery.FieldByName("TABTIPOPF").Value = qBuscaDadosPlano.FieldByName("TABTIPOPF").Value
    CurrentQuery.FieldByName("TIPOCONTAGEM").Value = qBuscaDadosPlano.FieldByName("TIPOCONTAGEM").Value
    CurrentQuery.FieldByName("TIPOPERIODO").Value = qBuscaDadosPlano.FieldByName("TIPOPERIODO").Value
  End If

  Set qBuscaDadosPlano = Nothing
End Sub
