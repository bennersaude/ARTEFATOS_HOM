'HASH: 482EBBC2EEC14D3F5D18C348B42E094C
'Macro: SAM_BENEFICIARIO_PFEVENTO
' Mauricio Ibelli -sms 2075 -27/03/2001 -Inibido condicao para considerar qdo e tipo de pf
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABELAPFEVENTO_OnChange()
  Dim qBuscaDadosPlano As Object
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

    If (qBuscaDadosPlano.FieldByName("TABPADRAOCONTAGEM").IsNull) Then
      CurrentQuery.FieldByName("TABPADRAOCONTAGEM").Value = 1
    Else
      CurrentQuery.FieldByName("TABPADRAOCONTAGEM").Value = qBuscaDadosPlano.FieldByName("TABPADRAOCONTAGEM").Value
    End If

    CurrentQuery.FieldByName("TABTIPOPF").AsInteger = qBuscaDadosPlano.FieldByName("TABTIPOPF").AsInteger
    CurrentQuery.FieldByName("CODIGOPF").Value = qBuscaDadosPlano.FieldByName("CODIGOPF").Value
    CurrentQuery.FieldByName("PERIODO").Value = qBuscaDadosPlano.FieldByName("PERIODO").Value
    CurrentQuery.FieldByName("TIPOPERIODO").Value = qBuscaDadosPlano.FieldByName("TIPOPERIODO").Value
  End If

  Set qBuscaDadosPlano = Nothing
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
  Dim QAUX As Object
  Set QAUX = NewQuery
  QAUX.Active = False
  QAUX.Clear
  QAUX.Add("SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEFICIARIO")
  QAUX.ParamByName("HBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  QAUX.Active = True

  If WebMode Then
    REGRACONTRATO.WebLocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  ElseIf VisibleMode Then
  	REGRACONTRATO.LocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  End If

  Set QAUX = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  Dim QAUX As Object
  Set QAUX = NewQuery
  QAUX.Active = False
  QAUX.Clear
  QAUX.Add("SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEFICIARIO")
  QAUX.ParamByName("HBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  QAUX.Active = True

  If WebMode Then
    REGRACONTRATO.WebLocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  ElseIf VisibleMode Then
  	REGRACONTRATO.LocalWhere = "CONTRATO = " + QAUX.FieldByName("CONTRATO").AsString
  End If

  Set QAUX = Nothing


    'TABELAPFEVENTO_OnPopup
  Dim qBeneficiario           As Object
  Dim vsNvl                   As String
  Dim vsCondicao              As String
  Dim vDataInfinita			  As String


  Set qBeneficiario = NewQuery

  'Selecionar as datas de adesão e cancelamento do beneficiário e formatá-las de acordo com o banco de dados.
  qBeneficiario.Clear
  qBeneficiario.Add("SELECT DATAADESAO,      ")
  qBeneficiario.Add("       DATACANCELAMENTO,")
  qBeneficiario.Add("		ATENDIMENTOATE   ")
  qBeneficiario.Add("  FROM SAM_BENEFICIARIO ")
  qBeneficiario.Add(" WHERE HANDLE = :HANDLE ")
  qBeneficiario.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qBeneficiario.Active = True

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

  'Se o beneficiário possui data de cancelamento ou atendimento até,
  'o filtro ficará restrito àqueles registros onde a data inicial é menor ou igual ao cancelamento ou atendimento até(se possuir) do beneficiário.
  If (Not qBeneficiario.FieldByName("ATENDIMENTOATE").IsNull) Then
    vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(qBeneficiario.FieldByName("ATENDIMENTOATE").AsDateTime)
  Else
    If (Not qBeneficiario.FieldByName("DATACANCELAMENTO").IsNull) Then
      vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(qBeneficiario.FieldByName("DATACANCELAMENTO").AsDateTime)
    End If
  End If

  Set qBeneficiario = Nothing

  'Filtrar também pela vigência do próprio registro que está sendo incluído/alterado.
  vsCondicao = vsCondicao + " AND DATAINICIAL <= " + SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  vsCondicao = vsCondicao + " AND " + vsNvl + "(DATAFINAL, " + vDataInfinita + ") >= " + SQLDate(CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)

  'As verificações feitas acima também serão feitas antes de gravar o registro, pois o usuário pode selecionar uma tabela PF evento e depois alterar as datas da vigência.

  vsCondicao = vsCondicao + " AND REGRA IS NULL"
  If WebMode Then
    TABELAPFEVENTO.WebLocalWhere = vsCondicao
  ElseIf VisibleMode Then
  	TABELAPFEVENTO.LocalWhere = vsCondicao
  End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.State = 3 Then
    If CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 1 Then
      CurrentQuery.FieldByName("PERIODO").Value = Null
    Else
      CurrentQuery.FieldByName("CODIGOPF").Value = Null
    End If
  End If


  Dim viResultado   As Integer
  Dim qBeneficiario As Object

  Set qBeneficiario = NewQuery

  qBeneficiario.Clear
  qBeneficiario.Add("SELECT DATAADESAO,      ")
  qBeneficiario.Add("       DATACANCELAMENTO,")
  qBeneficiario.Add("       ATENDIMENTOATE   ")
  qBeneficiario.Add("  FROM SAM_BENEFICIARIO ")
  qBeneficiario.Add(" WHERE HANDLE = :HANDLE ")
  qBeneficiario.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  qBeneficiario.Active = True

  viResultado = VerificaVigencia(qBeneficiario, 1)

  Set qBeneficiario = Nothing

  If (viResultado < 0) Then
    If (viResultado = -1) Then
      If (VisibleMode) Then
        MsgBox("A data inicial não pode ser anterior à adesão do beneficiário.")
      Else
        CancelDescription = "A data inicial não pode ser anterior à adesão do beneficiário."
      End If
    ElseIf (viResultado = -2) Then
      If (VisibleMode) Then
        MsgBox("A data final não pode ficar em aberto, pois o beneficiário possui data de cancelamento.")
      Else
        CancelDescription = "A data final não pode ficar em aberto, pois o beneficiário possui data de cancelamento."
      End If
    ElseIf (viResultado = -3) Then
      If (VisibleMode) Then
        MsgBox("A data final não pode ser posterior ao cancelamento do beneficiário.")
      Else
        CancelDescription = "A data final não pode ser posterior ao cancelamento do beneficiário."
      End If
    End If
    CanContinue = False
    Exit Sub
  End If

  Dim qContratoPFEvento As Object
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
      If (VisibleMode) Then
        MsgBox("A data inicial não pode ser anterior ao início da vigência do grupo de participação financeira.")
      Else
        CancelDescription = "A data inicial não pode ser anterior ao início da vigência do grupo de participação financeira."
      End If
    ElseIf (viResultado = -2) Then
      If (VisibleMode) Then
        MsgBox("A data final não pode ficar em aberto, pois o grupo de participação financeira possui vigência fechada.")
      Else
        CancelDescription = "A data final não pode ficar em aberto, pois o grupo de participação financeira possui vigência fechada."
      End If
    ElseIf (viResultado = -3) Then
      If (VisibleMode) Then
        MsgBox("A data final não pode ser posterior ao término da vigência do grupo de participação financeira.")
      Else
        CancelDescription = "A data final não pode ser posterior ao término da vigência do grupo de participação financeira."
      End If
    End If
    CanContinue = False
    Exit Sub
  End If

  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  If CurrentQuery.FieldByName("TABTIPOPF").AsInteger <> 3 Then
    Condicao = "AND ((TABELAPFEVENTO = " + CurrentQuery.FieldByName("TABELAPFEVENTO").AsString + ") OR (REGRACONTRATO IS NOT NULL)) "
    If (CurrentQuery.FieldByName("TABELAPFEVENTO").IsNull) Then ' É obrigatório nesses casos informar o grupo de PF
      CanContinue = False
      TABELAPFEVENTO.SetFocus
      bsShowMessage("Obrigatório informar o Grupo de participação financeira.", "E")
      Exit Sub
    End If
  Else
    Condicao = " AND REGRACONTRATO = " + CurrentQuery.FieldByName("REGRACONTRATO").AsString
  End If

  'Verificar se a vigência é compatível com a vigência do grupo de PF.
  If (Not CurrentQuery.FieldByName("TABELAPFEVENTO").IsNull) Then
    Dim qGrupoPF As Object
    Set qGrupoPF = NewQuery

    qGrupoPF.Clear
    qGrupoPF.Add("SELECT COUNT(1) QTD                                        ")
    qGrupoPF.Add("  FROM SAM_CONTRATO_PFEVENTO                               ")
    qGrupoPF.Add(" WHERE HANDLE = :HANDLE                                    ")
    qGrupoPF.Add("   AND (DATAFINAL IS NOT NULL AND DATAFINAL < :DATAINICIAL)")
    If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
      qGrupoPF.Add("    OR (DATAINICIAL > :DATAFINAL)")
      qGrupoPF.ParamByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
    End If
    qGrupoPF.ParamByName("HANDLE"     ).AsInteger  = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
    qGrupoPF.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
    qGrupoPF.Active = True

    If (qGrupoPF.FieldByName("QTD").AsInteger = 1) Then
      bsShowMessage("O grupo de participação financeira não está vigente no período especificado.", "E")
      CanContinue = False
      Set qGrupoPF = Nothing
      Exit Sub
    End If
    Set qGrupoPF = Nothing
  End If

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_BENEFICIARIO_PFEVENTO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "BENEFICIARIO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
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
           "Data final resultante : " + Format(DateAdd("m", CurrentQuery.FieldByName("PERIODO").AsInteger, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)), "E")
    DATAFINAL.SetFocus
    Exit Sub
  End If

  If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull And CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 2) Then ' Data final preenchida e tipo Variável
    Dim qBuscaContagens As Object
    Dim qAlteraVigencias As Object

    Set qBuscaContagens = NewQuery
    Set qAlteraVigencias = NewQuery

    qBuscaContagens.Active = False
    qBuscaContagens.Clear

    qAlteraVigencias.Active = False
    qAlteraVigencias.Clear

    ' Busca as contagens já efetuadas no contrato
    qBuscaContagens.Add("SELECT * FROM SAM_BENEFICIARIO_CONTPF WHERE BENEFICIARIO = :pBENEFICIARIO AND TABPF = (SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pTABELAPF) AND ((DATAFINAL > :pDATAFINAL) OR (DATAFINAL IS NULL))")
    qAlteraVigencias.Add("UPDATE SAM_BENEFICIARIO_CONTPF SET DATAFINAL = :pDATA WHERE BENEFICIARIO = :pBENEFICIARIO AND TABPF = (SELECT TABELAPFEVENTO FROM SAM_CONTRATO_PFEVENTO WHERE HANDLE = :pTABELAPF) AND ((DATAFINAL > :pDATA) OR (DATAFINAL IS NULL))")

    qBuscaContagens.ParamByName("pBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    qBuscaContagens.ParamByName("pTABELAPF").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
    qBuscaContagens.ParamByName("pDATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

    qAlteraVigencias.ParamByName("pBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    qAlteraVigencias.ParamByName("pTABELAPF").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
    qAlteraVigencias.ParamByName("pDATA").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

    qBuscaContagens.Active = True
    ' Há contagens com a data final maior que a data final da regra
		If (Not qBuscaContagens.EOF) Then
	  		If WebMode Then
    			qAlteraVigencias.ExecSQL
    			bsShowMessage("Existiam contagens vigentes para esta regra. As contagens foram atualizadas com a data final informada nesta regra!", "I")
  			ElseIf VisibleMode Then
	    		If MsgBox("Existem contagens vigentes para esta regra, deseja continuar" + Chr(13) + "e atualizar a vigência dessas contagens com esta data final ?", vbYesNo) = vbYes Then
      			qAlteraVigencias.ExecSQL
    			Else
      			CanContinue = False
      			DATAFINAL.SetFocus
      			Exit Sub
    			End If
  			End If
		End If

		Set qBuscaContagens = Nothing
		Set qAlteraVigencias = Nothing
  End If

End Sub

Public Sub TABTIPOPF_OnChange()
  If CurrentQuery.State = 3 Then
    If CurrentQuery.FieldByName("TABTIPOPF").AsInteger = 1 Then
      CurrentQuery.FieldByName("PERIODO").Value = Null
    Else
      CurrentQuery.FieldByName("CODIGOPF").Value = Null
    End If
  End If
End Sub

'Retorna:
'    1 : caso a vigência seja compatível;
'   -1 : caso a data inicial da vigência do registro seja anterior ao início da outra vigência relacionada;
'   -2 : caso a outra vigência relacionada tenha data final e a data final da vigência do registro esteja em aberto;
'   -3 : caso a data final da vigência do registro seja posterior ao término da outra vigência relacioanda.
Function VerificaVigencia(Query As Object, Tipo As Integer) As Integer
  Dim vsInicio As String
  Dim vsFinal  As String

  If (Tipo = 1) Then
    vsInicio = "DATAADESAO"

    If (Query.FieldByName("ATENDIMENTOATE").IsNull) Then
		vsFinal  = "DATACANCELAMENTO"
	Else
		vsFinal  = "ATENDIMENTOATE"
    End If

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

