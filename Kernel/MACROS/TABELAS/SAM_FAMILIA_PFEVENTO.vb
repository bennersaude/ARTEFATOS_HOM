'HASH: 5FB29B724AE995901E711AD9713549D1
'Macro: SAM_FAMILIA_PFEVENTO
' Mauricio Ibelli -sms 2075 -27/03/2001 -Inibido condicao para considerar qdo e tipo de pf
'#Uses "*bsShowMessage"

Option Explicit
Dim bIntercambiavel	As Boolean

Public Sub DATAINICIAL_OnExit()
  AlterarCondicaoPFevento()
End Sub

Public Sub INTERCAMBIAVEL_OnChange()
	bIntercambiavel = Not bIntercambiavel
End Sub


Public Sub TABELAPFEVENTO_OnChange()
  Dim qBuscaDadosPlano As BPesquisa
  Set qBuscaDadosPlano = NewQuery

  qBuscaDadosPlano.Clear
  qBuscaDadosPlano.Add("SELECT * FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pPFEVENTO")
  qBuscaDadosPlano.ParamByName("pPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
  qBuscaDadosPlano.Active = True

  If (Not qBuscaDadosPlano.EOF) Then
    If (qBuscaDadosPlano.FieldByName("ACEITAFINANCIAMENTO").IsNull) Then
      CurrentQuery.FieldByName("ACEITAFINANCIAMENTO").Value = "N"
    Else
      CurrentQuery.FieldByName("ACEITAFINANCIAMENTO").Value = qBuscaDadosPlano.FieldByName("ACEITAFINANCIAMENTO").Value
    End If

    If (qBuscaDadosPlano.FieldByName("ACEITAPARCELAMENTO").IsNull) Then
      CurrentQuery.FieldByName("ACEITAPARCELAMENTO").Value = "N"
    Else
      CurrentQuery.FieldByName("ACEITAPARCELAMENTO").Value = qBuscaDadosPlano.FieldByName("ACEITAPARCELAMENTO").Value
    End If

    If (qBuscaDadosPlano.FieldByName("INTERCAMBIAVEL").IsNull) Then
      CurrentQuery.FieldByName("INTERCAMBIAVEL").Value = "N"
    Else
      CurrentQuery.FieldByName("INTERCAMBIAVEL").Value = qBuscaDadosPlano.FieldByName("INTERCAMBIAVEL").Value
    End If

    If (qBuscaDadosPlano.FieldByName("TABPADRAOCONTAGEM").IsNull) Then
      CurrentQuery.FieldByName("TABPADRAOCONTAGEM").Value = 1
    Else
      CurrentQuery.FieldByName("TABPADRAOCONTAGEM").Value = qBuscaDadosPlano.FieldByName("TABPADRAOCONTAGEM").Value
    End If

    CurrentQuery.FieldByName("TABTIPOPF").AsInteger = qBuscaDadosPlano.FieldByName("TABTIPOPF").AsInteger
    CurrentQuery.FieldByName("CODIGOPF").Value = qBuscaDadosPlano.FieldByName("CODIGOPF").Value
    CurrentQuery.FieldByName("PERIODO").Value = qBuscaDadosPlano.FieldByName("PERIODO").Value
    If (qBuscaDadosPlano.FieldByName("TIPOCONTAGEM").AsString = "C") Then
      CurrentQuery.FieldByName("TIPOCONTAGEM").Value = "F"
    Else
      CurrentQuery.FieldByName("TIPOCONTAGEM").Value = qBuscaDadosPlano.FieldByName("TIPOCONTAGEM").Value
    End If
    CurrentQuery.FieldByName("TIPOPERIODO").Value = qBuscaDadosPlano.FieldByName("TIPOPERIODO").Value
  End If

  Set qBuscaDadosPlano = Nothing
End Sub

Public Sub TABELAPFEVENTO_OnPopup(ShowPopup As Boolean)

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
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim QAUX As BPesquisa
  Set QAUX = NewQuery
  QAUX.Active = False
  QAUX.Clear
  QAUX.Add("SELECT CONTRATO FROM SAM_FAMILIA WHERE HANDLE = :HFAMILIA")
  QAUX.ParamByName("HFAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  QAUX.Active = True

  If WebMode Then
  	REGRACONTRATO.WebLocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  ElseIf VisibleMode Then
  	REGRACONTRATO.LocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  End If


  Set QAUX = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim QAUX As BPesquisa

  AlterarCondicaoPFevento()

  Set QAUX = NewQuery
  QAUX.Active = False
  QAUX.Clear
  QAUX.Add("SELECT CONTRATO FROM SAM_FAMILIA WHERE HANDLE = :HFAMILIA")
  QAUX.ParamByName("HFAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  QAUX.Active = True

  If WebMode Then
  	REGRACONTRATO.WebLocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  ElseIf VisibleMode Then
  	REGRACONTRATO.LocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  End If

  Set QAUX = Nothing
End Sub


Function AlterarCondicaoPFevento()
  Dim vsCondicao              As String
  Dim vDataInfinita			  As String
  Dim vsNvl                   As String
  Dim qFamilia                As BPesquisa
  Set qFamilia = NewQuery

  'Selecionar as datas de adesão e cancelamento do beneficiário e formatá-las de acordo com o banco de dados.
  qFamilia.Clear
  qFamilia.Add("SELECT DATAADESAO,     ")
  qFamilia.Add("       DATACANCELAMENTO")
  qFamilia.Add("  FROM SAM_FAMILIA     ")
  qFamilia.Add(" WHERE HANDLE = :HANDLE")
  qFamilia.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  qFamilia.Active = True


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
  vsCondicao = vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= " + SQLDate(qFamilia.FieldByName("DATAADESAO").AsDateTime)
  If (Not qFamilia.FieldByName("DATACANCELAMENTO").IsNull) Then
    'Se o beneficiário possui data de cancelamento, o filtro ficará restrito àqueles registros onde a data inicial é menor ou igual ao cancelamento do beneficiário.
    vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(qFamilia.FieldByName("DATACANCELAMENTO").AsDateTime)
  End If
  Set qFamilia = Nothing

  'Filtrar também pela vigência do próprio registro que está sendo incluído/alterado.
  If WebMode Then
  	vsCondicao = vsCondicao + " AND DATAINICIAL <= @CAMPO(DATAINICIAL)"
  	vsCondicao = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= @CAMPO(DATAINICIAL)"
  ElseIf VisibleMode Then
  	vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  	vsCondicao = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >=  " + SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  End If

  'As verificações feitas acima também serão feitas antes de gravar o registro, pois o usuário pode selecionar uma tabela PF evento e depois alterar as datas da vigência.
  vsCondicao = vsCondicao + " AND REGRA IS NULL"

  If WebMode Then
	TABELAPFEVENTO.WebLocalWhere = vsCondicao
  ElseIf VisibleMode Then
  	TABELAPFEVENTO.LocalWhere = vsCondicao
  End If

End Function


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qFamilia As BPesquisa
  Set qFamilia = NewQuery
  Dim viResultado As Long

  qFamilia.Clear
  qFamilia.Add("SELECT DATAADESAO,     ")
  qFamilia.Add("       DATACANCELAMENTO")
  qFamilia.Add("  FROM SAM_FAMILIA     ")
  qFamilia.Add(" WHERE HANDLE = :HANDLE")
  qFamilia.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  qFamilia.Active = True

  viResultado = VerificaVigencia(qFamilia, 1)

  Set qFamilia = Nothing

  If (viResultado < 0) Then
    If (viResultado = -1) Then
        bsShowMessage("A data inicial não pode ser anterior à adesão da família.", "E")
    ElseIf (viResultado = -2) Then
        bsShowMessage("A data final não pode ficar em aberto, pois a família possui data de cancelamento.", "E")
    ElseIf (viResultado = -3) Then
        bsShowMessage("A data final não pode ser posterior ao cancelamento da família.", "E")
    End If
    CanContinue = False
    Exit Sub
  End If

  Dim qContratoPFEvento As BPesquisa
  Set qContratoPFEvento = NewQuery

  qContratoPFEvento.Clear
  qContratoPFEvento.Add("SELECT DATAINICIAL,         ")
  qContratoPFEvento.Add("       DATAFINAL            ")
  qContratoPFEvento.Add("  FROM SAM_CONTRATO_PFEVENTO")
  qContratoPFEvento.Add(" WHERE HANDLE = :HANDLE     ")
  qContratoPFEvento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
  qContratoPFEvento.Active = True

  viResultado = VerificaVigencia(qContratoPFEvento, 2)

  Set qContratoPFEvento = Nothing

  If (viResultado < 0) Then
    If (viResultado = -1) Then
        bsShowMessage("A data inicial não pode ser anterior ao início da vigência do grupo de participação financeira.", "E")
    ElseIf (viResultado = -2) Then
        bsShowMessage("A data final não pode ficar em aberto, pois o grupo de participação financeira possui vigência fechada.", "E")
    ElseIf (viResultado = -3) Then
        bsShowMessage("A data final não pode ser posterior ao término da vigência do grupo de participação financeira.", "E")
    End If
    CanContinue = False
    Exit Sub
  End If

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  'Balani SMS 4595 18/08/2005
  If (CurrentQuery.FieldByName("TABTIPOPF").AsInteger <> 3) Then
    Condicao = " AND ((TABELAPFEVENTO = " + CurrentQuery.FieldByName("TABELAPFEVENTO").AsString + ") OR (REGRACONTRATO IS NOT NULL))"
    'Condicao =Condicao +" AND TABTIPOPF = " +CurrentQuery.FieldByName("TABTIPOPF").AsString
    If (CurrentQuery.FieldByName("TABELAPFEVENTO").IsNull) Then ' É obrigatório nesses casos informar o grupo de PF
      CanContinue = False
      TABELAPFEVENTO.SetFocus
      bsShowMessage("Obrigatório informar o ""Grupo de participação financeira.""", "E")
      Exit Sub
    End If
  Else
    Condicao = ""
  End If
  'final SMS 4595
  'Condicao =" AND TABELAPFEVENTO = " +CurrentQuery.FieldByName("TABELAPFEVENTO").AsString
  'Condicao =Condicao +" AND TABTIPOPF = " +CurrentQuery.FieldByName("TABTIPOPF").AsString
  'If CurrentQuery.FieldByName("TABTIPOPF").AsInteger =2 Then
  '  Condicao =Condicao +" AND TIPOCONTAGEM = '" +CurrentQuery.FieldByName("TIPOCONTAGEM").AsString +"'"
  'End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_PFEVENTO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FAMILIA", Condicao)

  If (Linha = "") Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

  'SMS 40817 - Anderson Lonardoni - 03/05/2005
  'SE O TIPO DE CONTAGEM FOR NA FAMÍLIA O TIPO DE PERÍODO DEVERÁ SER CIVIL, POR ADESÃO DO CONTRATO OU DA FAMÍLIA
  If (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "F") And (CurrentQuery.FieldByName("TIPOPERIODO").AsString = "B") Then
    CanContinue = False
    bsShowMessage("Tipo de contagem na família exige que o tipo de período seja civil, por adesão do contrato ou da família!", "E")
    Exit Sub
  End If

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
           "Data final resultante : " + Format(DateAdd("m", CurrentQuery.FieldByName("PERIODO").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)), "I")
    DATAFINAL.SetFocus
    Exit Sub
  End If

  If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull And CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 2) Then ' Data final preenchida e tipo Variável
    Dim qBuscaContagens As BPesquisa
    Dim qAlteraVigencias As BPesquisa

    Set qBuscaContagens = NewQuery
    Set qAlteraVigencias = NewQuery

    qBuscaContagens.Active = False
    qBuscaContagens.Clear

    qAlteraVigencias.Active = False
    qAlteraVigencias.Clear

    ' Busca as contagens já efetuadas no contrato
    If (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "F") Then ' Contagem na família
      qBuscaContagens.Add("SELECT * FROM SAM_FAMILIA_CONTPF WHERE FAMILIA = :pFAMILIA AND TABPF = (SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pTABELAPF) AND ((DATAFINAL > :pDATAFINAL) OR (DATAFINAL IS NULL))")
      qAlteraVigencias.Add("UPDATE SAM_FAMILIA_CONTPF SET DATAFINAL = :pDATA WHERE FAMILIA = :pFAMILIA AND TABPF = (SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pTABELAPF) AND ((DATAFINAL > :pDATA) OR (DATAFINAL IS NULL))")
    ElseIf (CurrentQuery.FieldByName("TIPOCONTAGEM").AsString = "B") Then ' Contagem no beneficiário
      qBuscaContagens.Add("SELECT * FROM SAM_BENEFICIARIO_CONTPF WHERE BENEFICIARIO IN (SELECT DISTINCT HANDLE FROM SAM_BENEFICIARIO WHERE FAMILIA = :pFAMILIA) AND TABPF = (SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pTABELAPF) AND ((DATAFINAL > :pDATAFINAL) OR (DATAFINAL IS NULL))")
      qAlteraVigencias.Add("UPDATE SAM_BENEFICIARIO_CONTPF SET DATAFINAL = :pDATA WHERE BENEFICIARIO IN (SELECT DISTINCT HANDLE FROM SAM_BENEFICIARIO WHERE FAMILIA = :pFAMILIA) AND TABPF = (SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pTABELAPF) AND ((DATAFINAL > :pDATA) OR (DATAFINAL IS NULL))")
    End If
    qBuscaContagens.ParamByName("pFAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
    qBuscaContagens.ParamByName("pTABELAPF").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
    qBuscaContagens.ParamByName("pDATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

    qAlteraVigencias.ParamByName("pFAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
    qAlteraVigencias.ParamByName("pTABELAPF").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
    qAlteraVigencias.ParamByName("pDATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

    qBuscaContagens.Active = True
    ' Há contagens com a data final maior que a data final da regra
    If (Not qBuscaContagens.EOF) Then
      bsShowMessage("Vigêngia atualizada com esta data final", "I")
        qAlteraVigencias.ExecSQL
    End If
    Set qBuscaContagens = Nothing
    Set qAlteraVigencias = Nothing
  End If

	If (bIntercambiavel And CurrentQuery.FieldByName("TIPOCONTAGEM").AsString <> "F") Then
		bsShowMessage("Para PFs intercambiáveis, a contagem deve ser feita na 'Família'.", "E")
		CanContinue = False
		Exit Sub
	End If

End Sub

'Retorna:
'    1 : caso a vigência seja compatível;
'   -1 : caso a data inicial da vigência do registro seja anterior ao início da outra vigência relacionada;
'   -2 : caso a outra vigência relacionada tenha data final e a data final da vigência do registro esteja em aberto;
'   -3 : caso a data final da vigência do registro seja posterior ao término da outra vigência relacioanda.
Function VerificaVigencia(Query As BPesquisa, Tipo As Integer) As Integer
  Dim vsInicio As String
  Dim vsFinal  As String

  If (Tipo = 1) Then
    vsInicio = "DATAADESAO"
    vsFinal  = "DATACANCELAMENTO"
  ElseIf (Tipo = 2) Then
    vsInicio = "DATAINICIAL"
    vsFinal  = "DATAFINAL"
  End If

  VerificaVigencia = 1
  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < Query.FieldByName(vsInicio).AsDateTime) Then
    VerificaVigencia = -1
  Else
    If (Not Query.FieldByName(vsFinal).IsNull) Then
      If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
        VerificaVigencia = -2
      Else
        If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime > Query.FieldByName(vsFinal).AsDateTime) Then
          VerificaVigencia = -3
        End If
      End If
    End If
  End If
End Function

