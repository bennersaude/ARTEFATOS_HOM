'HASH: D0516660461497A47FB89E3244A6C399
'Macro: SAM_PLANO_LIMITACAO
'#Uses "*bsShowMessage"
'#Uses "*TipoPeriodoLimiteValido"

'Última atualização: 24/04/2002
' Milton -SMS 8763

Public Sub BOTAOATUALIZACONTRATO_OnClick()

	ATUALIZARCONTRATO(True)

End Sub


Public Sub ATUALIZARCONTRATO(CanContinue As Boolean)

  If CurrentQuery.FieldByName("LIMITACAO").IsNull Then
    bsShowMessage("Limitação sem valor!", "I")
    Exit Sub
  End If


  Dim q1 As Object
  Set q1 = NewQuery
  Dim q2 As Object
  Set q2 = NewQuery
  Dim qlimitacaoFXeMOD As Object
  Set qlimitacaoFXeMOD = NewQuery
  Dim qinsertlimitacaoFXeMOD As Object
  Set qinsertlimitacaoFXeMOD = NewQuery
  Dim qlimitacaoMOD As Object
  Set qlimitacaoMOD = NewQuery
  Dim qinsertlimitacaoMOD As Object
  Set qinsertlimitacaoMOD = NewQuery

  Dim DataInicial As Date
  Dim DataFinal As Date

  Dim q3 As Object
  Set q3 = NewQuery
  Dim vHContratoLimitacao As Long

  Dim HandleFiltro As Integer
  Dim Filtro As Object
  Set Filtro = CreateBennerObject("SamFiltro.Filtro")

inicio :
  If (VisibleMode) Then
    HandleFiltro = Filtro.Exec(CurrentSystem, CurrentUser, 800, "DATAINICIAL.ob|DATAFINAL", "Atualizar Contratos")

    q1.Add("SELECT DATAINICIAL, DATAFINAL    ")
    q1.Add("FROM RF_FILTRO                   ")
    q1.Add("WHERE HANDLE=:PHANDLEFILTRO      ")
    q1.ParamByName("PHANDLEFILTRO").Value = HandleFiltro
    q1.Active = True

    If(Not q1.FieldByName("DATAFINAL").IsNull)And(q1.FieldByName("DATAFINAL").AsDateTime <q1.FieldByName("DATAINICIAL").AsDateTime)Then
	  bsShowMessage("Data final não pode ser anterior à data inicial.", "I")
	  GoTo inicio
	Else
	  DataInicial = q1.FieldByName("DATAINICIAL").AsDateTime
	  DataFinal = q1.FieldByName("DATAFINAL").AsDateTime
    End If

    If (q1.FieldByName("DATAINICIAL").IsNull) Then
      bsShowMessage("Data inicial não informada, atualização não efetuada", "E")
      Set q1 = Nothing
      Set q2 = Nothing
      Set q3 = Nothing
      Set qlimitacaoFXeMOD = Nothing
      Set qinsertlimitacaoFXeMOD = Nothing
      Set qlimitacaoMOD = Nothing
      Set qinsertlimitacaoMOD = Nothing
      Exit Sub
    End If
  ElseIf (WebMode) Then
      If(Not CurrentVirtualQuery.FieldByName("DATAFINAL").IsNull) And (CurrentVirtualQuery.FieldByName("DATAFINAL").AsDateTime < CurrentVirtualQuery.FieldByName("DATAINICIAL").AsDateTime)Then
	    bsShowMessage("Data final não pode ser anterior à data inicial.", "E")
	    CanContinue = False
	    Exit Sub
	  Else
	    DataInicial = CurrentVirtualQuery.FieldByName("DATAINICIAL").AsDateTime
	    DataFinal = CurrentVirtualQuery.FieldByName("DATAFINAL").AsDateTime
	  End If
  End If

q2.Clear
q2.Add("SELECT HANDLE CONTRATO ")
q2.Add("  FROM SAM_CONTRATO C  ")
q2.Add(" WHERE C.DATACANCELAMENTO IS NULL")
q2.Add("   AND C.PLANO=:pPLANO")
q2.Add("   and handle not in (SELECT CONTRATO  ")
q2.Add("                        FROM SAM_CONTRATO_LIMITACAO ")
q2.Add("                       WHERE LIMITACAO=:pLIMITACAO)")

q2.ParamByName("pPLANO").Value = CurrentQuery.FieldByName("PLANO").Value
q2.ParamByName("pLIMITACAO").Value = CurrentQuery.FieldByName("LIMITACAO").Value
q2.Active = True

If Not InTransaction Then
  StartTransaction
End If

q3.Clear
q3.Add("INSERT INTO SAM_CONTRATO_LIMITACAO																")
q3.Add("(HANDLE, TIPOCONTAGEM, LIMITACAO, PERIODO, CONTRATO, DATAINICIAL, DATAFINAL, QTDLIMITE, INTERCAMBIAVEL, PLANO, TIPOLIMITACAO, TIPOPERIODO, TABTIPOLIMITE, TABTIPOVALOR, VLRLIMITE, TABELAUS )	")
q3.Add("VALUES																							")
q3.Add("(:pHANDLE, :pTIPOCONTAGEM, :pLIMITACAO, :pPERIODO, :pCONTRATO, :pDATAINICIAL, :pDATAFINAL, :pQTDLIMITE, :pINTERCAMBIAVEL, :pPLANO, :pTIPOLIMITACAO, :PTIPOPERIODO, :pTABTIPOLIMITE, :pTABTIPOVALOR, :pVLRLIMITE, :pTABELAUS )")

qlimitacaoFXeMOD.Clear
qlimitacaoFXeMOD.Add("SELECT C.*, A.HANDLE CONTRATO ")
qlimitacaoFXeMOD.Add("  FROM SAM_PLANO_LIMITACAO_FX C, SAM_PLANO_LIMITACAO B , SAM_CONTRATO A  ")
qlimitacaoFXeMOD.Add(" WHERE C.PLANOLIMITACAO = :pPLANOLIMITACAO ")
qlimitacaoFXeMOD.Add("   AND B.HANDLE         = C.PLANOLIMITACAO ")
qlimitacaoFXeMOD.Add("   AND A.PLANO          = B.PLANO          ")
qlimitacaoFXeMOD.Add("   AND A.HANDLE         = :pContrato       ")


qlimitacaoMOD.Add("SELECT A.*, D.HANDLE CONTRATOMOD     ")
qlimitacaoMOD.Add("  FROM SAM_PLANO_LIMITACAO_MOD A,   ")
qlimitacaoMOD.Add("       SAM_PLANO_MOD B,             ")
qlimitacaoMOD.Add("       SAM_MODULO C,                ")
qlimitacaoMOD.Add("       SAM_CONTRATO_MOD D           ")
qlimitacaoMOD.Add(" WHERE A.PLANOLIMITACAO = :pPLANOLIMITACAO ")
qlimitacaoMOD.Add("   And B.HANDLE = A.PLANOMODULO     ")
qlimitacaoMOD.Add("   And C.HANDLE = B.MODULO          ")
qlimitacaoMOD.Add("   And C.HANDLE = D.MODULO          ")
qlimitacaoMOD.Add("   And D.CONTRATO = :pContrato      ")


qinsertlimitacaoFXeMOD.Add("INSERT INTO SAM_CONTRATO_LIMITACAO_FX ")
qinsertlimitacaoFXeMOD.Add("        ( HANDLE, CONTRATO, CONTRATOLIMITACAO, QUANTIDADE, NIVELAUTORIZACAO, LIMITACAO ) ")
qinsertlimitacaoFXeMOD.Add(" VALUES (:HANDLE,:CONTRATO,:CONTRATOLIMITACAO,:QUANTIDADE,:NIVELAUTORIZACAO,:LIMITACAO ) ")


qinsertlimitacaoMOD.Add("INSERT INTO SAM_CONTRATO_LIMITACAO_MOD ")
qinsertlimitacaoMOD.Add("        ( HANDLE, CONTRATOLIMITACAO ,CONTRATO, CONTRATOMODULO ) ")
qinsertlimitacaoMOD.Add(" VALUES (:HANDLE,:CONTRATOLIMITACAO,:CONTRATO,:CONTRATOMODULO ) ")


On Error GoTo FIM


While Not q2.EOF
  ' Inserindo a limitação no contrato
  vHContratoLimitacao = NewHandle("SAM_CONTRATO_LIMITACAO")
  q3.ParamByName("pHANDLE" ).Value = vHContratoLimitacao
  q3.ParamByName("pPLANO" ).Value = CurrentQuery.FieldByName("PLANO" ).Value
  q3.ParamByName("pTIPOCONTAGEM" ).Value = CurrentQuery.FieldByName("TIPOCONTAGEM").AsString
  q3.ParamByName("pLIMITACAO" ).Value = CurrentQuery.FieldByName("LIMITACAO" ).AsInteger
  q3.ParamByName("pPERIODO" ).Value = CurrentQuery.FieldByName("PERIODO" ).AsInteger
  q3.ParamByName("pCONTRATO" ).Value = q2.FieldByName("CONTRATO").AsInteger
  q3.ParamByName("pDATAINICIAL" ).Value = DataInicial
  q3.ParamByName("pQTDLIMITE" ).Value = CurrentQuery.FieldByName("QTDLIMITE" ).AsInteger
  q3.ParamByName("pINTERCAMBIAVEL").Value = CurrentQuery.FieldByName("INTERCAMBIAVEL").AsString
  q3.ParamByName("pTIPOLIMITACAO" ).Value = CurrentQuery.FieldByName("TIPOLIMITACAO" ).Value
  q3.ParamByName("pTIPOPERIODO" ).Value = CurrentQuery.FieldByName("TIPOPERIODO" ).Value
  q3.ParamByName("pTABTIPOLIMITE" ).AsInteger = CurrentQuery.FieldByName("TABTIPOLIMITE" ).AsInteger
  If CurrentQuery.FieldByName("TABTIPOLIMITE" ).AsInteger = 2 Then ' tipo limite "por valor"
    q3.ParamByName("pTABTIPOVALOR" ).Value = CurrentQuery.FieldByName("TABTIPOVALOR" ).Value
    q3.ParamByName("pVLRLIMITE" ).Value = CurrentQuery.FieldByName("VLRLIMITE" ).Value
    If CurrentQuery.FieldByName("TABTIPOVALOR").AsInteger = 2 Then 'Valor por US
      q3.ParamByName("pTABELAUS").Value = CurrentQuery.FieldByName("TABELAUS").Value
    Else
      q3.ParamByName("pTABELAUS").DataType = ftInteger
      q3.ParamByName("pTABELAUS").Clear
    End If
  Else
    q3.ParamByName("pTABTIPOVALOR" ).DataType = ftInteger
    q3.ParamByName("pTABTIPOVALOR" ).Clear
    q3.ParamByName("pVLRLIMITE" ).DataType = ftFloat
    q3.ParamByName("pVLRLIMITE" ).Clear
    q3.ParamByName("pTABELAUS" ).DataType = ftInteger
    q3.ParamByName("pTABELAUS" ).Clear
  End If

  If (VisibleMode) Then
    If (q1.FieldByName("DATAINICIAL").IsNull) Then
      q3.ParamByName("pDATAFINAL").DataType = ftDateTime
      q3.ParamByName("pDATAFINAL").Clear
    Else
      q3.ParamByName("pDATAFINAL").Value = DataFinal
    End If
  ElseIf (WebMode) Then
      If (CurrentVirtualQuery.FieldByName("DATAINICIAL").IsNull) Then
        q3.ParamByName("pDATAFINAL").DataType = ftDateTime
        q3.ParamByName("pDATAFINAL").Clear
      Else
        q3.ParamByName("pDATAFINAL").Value = DataFinal
      End If
  End If
  q3.ExecSQL

  ' Inserindo a limitação no contrato
  qlimitacaoFXeMOD.ParamByName("pPLANOLIMITACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qlimitacaoFXeMOD.ParamByName("pContrato" ).AsInteger = q2.FieldByName("CONTRATO").AsInteger
  qlimitacaoFXeMOD.Active = True

  While Not qlimitacaoFXeMOD.EOF
    qinsertlimitacaoFXeMOD.Active = False
    qinsertlimitacaoFXeMOD.ParamByName("HANDLE" ).Value = NewHandle("SAM_CONTRATO_LIMITACAO_FX")
    qinsertlimitacaoFXeMOD.ParamByName("CONTRATOLIMITACAO").AsInteger = vHContratoLimitacao
    qinsertlimitacaoFXeMOD.ParamByName("CONTRATO" ).AsInteger = q2.FieldByName("CONTRATO").AsInteger
    qinsertlimitacaoFXeMOD.ParamByName("LIMITACAO" ).AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
    qinsertlimitacaoFXeMOD.ParamByName("QUANTIDADE" ).AsFloat = qlimitacaoFXeMOD.FieldByName("QUANTIDADE" ).AsFloat
    qinsertlimitacaoFXeMOD.ParamByName("NIVELAUTORIZACAO" ).AsInteger = qlimitacaoFXeMOD.FieldByName("NIVELAUTORIZACAO").AsInteger
    qinsertlimitacaoFXeMOD.ExecSQL

    qlimitacaoFXeMOD.Next
  Wend



  qlimitacaoMOD.ParamByName("pPLANOLIMITACAO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qlimitacaoMOD.ParamByName("pContrato" ).AsInteger = q2.FieldByName("CONTRATO").AsInteger
  qlimitacaoMOD.Active = True

  While Not qlimitacaoMOD.EOF
    qinsertlimitacaoMOD.Active = False
    qinsertlimitacaoMOD.ParamByName("HANDLE" ).Value = NewHandle("SAM_CONTRATO_LIMITACAO_MOD")
    qinsertlimitacaoMOD.ParamByName("CONTRATOLIMITACAO").AsInteger = vHContratoLimitacao
    qinsertlimitacaoMOD.ParamByName("CONTRATO" ).AsInteger = q2.FieldByName("CONTRATO").AsInteger
    qinsertlimitacaoMOD.ParamByName("CONTRATOMODULO" ).AsInteger = qlimitacaoMOD.FieldByName("CONTRATOMOD").AsInteger
    qinsertlimitacaoMOD.ExecSQL

    qlimitacaoMOD.Next
  Wend


  q2.Next
Wend
Commit
Set q1 = Nothing
Set q2 = Nothing
Set q3 = Nothing
Set qlimitacaoFXeMOD = Nothing
Set qinsertlimitacaoFXeMOD = Nothing
Set qlimitacaoMOD = Nothing
Set qinsertlimitacaoMOD = Nothing

bsShowMessage("Atualização completa ", "I")
Exit Sub

FIM :
Rollback
bsShowMessage("Ocorreu o seguinte erro ao atualizar contrato(s)" + Str(Error), "E")
Set q1 = Nothing
Set q2 = Nothing
Set q3 = Nothing
Set qlimitacaoFXeMOD = Nothing
Set qinsertlimitacaoFXeMOD = Nothing
Set qlimitacaoMOD = Nothing
Set qinsertlimitacaoMOD = Nothing

End Sub

Public Sub LIMITACAO_OnChange()
  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODO.Visible = False
  Else
    PERIODO.Visible = True
  End If

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_AfterScroll()
  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODO.Visible = False
  Else
    PERIODO.Visible = True
  End If

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If(CurrentQuery.FieldByName("TIPOCONTAGEM").Value = "B")And _
     (CurrentQuery.FieldByName("INTERCAMBIAVEL").Value = "S")Then
    bsShowMessage("Acionar intercambiável somente para tipo de contagem contrato ou família", "E")
    CanContinue = False
  Else
    CanContinue = True
  End If

  CanContinue = TipoPeriodoLimiteValido(CurrentQuery.FieldByName("TIPOCONTAGEM").AsString, CurrentQuery.FieldByName("TIPOPERIODO").AsString)

  'SMS 61198 - Matheus - Início
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT PERIODICIDADE FROM SAM_LIMITACAO WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("LIMITACAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("PERIODICIDADE").AsInteger = 2 Then  CurrentQuery.FieldByName("PERIODO").AsInteger = 1

  Set SQL = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOATUALIZACONTRATO" Then
		ATUALIZARCONTRATO(CanContinue)
	End If
End Sub
