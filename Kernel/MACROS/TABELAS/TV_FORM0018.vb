'HASH: D45B0462F33C2E4BF8FB48182FD927FC
'#Uses "*bsShowMessage"


Public Sub BOTAOMOVFIN_OnClick()
	Dim qAux As Object

	Set qAux = NewQuery

	qAux.Add("SELECT HANDLE FROM R_RELATORIOS")
	qAux.Add(" WHERE CODIGO = 'BEN070'")
 	qAux.Active = True

	ReportPreview(qAux.FieldByName("HANDLE").AsInteger, "", False, False)

 	Set qAux = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	Dim vCriterio As String
	vCriterio = "HANDLE NOT IN (SELECT MOTIVOFALECIMENTOTITULAR FROM SAM_PARAMETROSBENEFICIARIO" + _
        " UNION" + _
        " SELECT MOTIVOMIGRACAO FROM SAM_PARAMETROSBENEFICIARIO" + _
        "  UNION" + _
    	" SELECT MIGRACAONREGREG FROM SAM_PARAMETROSBENEFICIARIO)"
    vCriterio = vCriterio + "AND HANDLE NOT IN (SELECT MOTIVOFALECIMENTO FROM SAM_PARAMETROSBENEFICIARIO)"

	If VisibleMode Then
		MOTIVO.LocalWhere = vCriterio
	ElseIf WebMode Then
		MOTIVO.WebLocalWhere = vCriterio
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object

  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    RecordHandleOfTable("SAM_FAMILIA"), _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido cancelar o módulo por motivo de suspensão!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set BSBen001Dll = Nothing


  If Not CurrentQuery.FieldByName("DATAATEND").IsNull Then
  	If CurrentQuery.FieldByName("DATAATEND").AsDateTime <= CurrentQuery.FieldByName("DATA").AsDateTime Then
  		bsShowMessage("A data de atendimento até deve ser posterior a data de cancelamento.","E")
  		CanContinue = False
		Exit Sub
	End If
  End If

  Set SQL = NewQuery
  SQL.Add("SELECT DATACANCELAMENTO")
  SQL.Add("  FROM SAM_FAMILIA_MOD    ")
  SQL.Add(" WHERE HANDLE = :HANDLE")

  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_FAMILIA_MOD")

  SQL.Active = True

  Dim Obj As Object
  Set Obj = CreateBennerObject("SAMCANCELAMENTO.Cancelar")

  If VisibleMode Then
	bsShowMessage(Obj.FamiliaModulo(CurrentSystem, RecordHandleOfTable("SAM_FAMILIA_MOD"),CurrentQuery.FieldByName("DATA").AsDateTime,CurrentQuery.FieldByName("MOTIVO").AsInteger, CurrentQuery.FieldByName("DATAATEND").AsDateTime), "I")
  ElseIf WebMode Then
  	If SQL.FieldByName("DATACANCELAMENTO").IsNull Then
		bsShowMessage(Obj.FamiliaModulo(CurrentSystem, RecordHandleOfTable("SAM_FAMILIA_MOD") ,CurrentQuery.FieldByName("DATA").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger, CurrentQuery.FieldByName("DATAATEND").AsDateTime),  "I")
 	Else
	 	bsShowMessage("Modulo da Familia já cancelado!","I")
 	End If
  End If

  Set Obj = Nothing



End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOMOVFIN" Then
		BOTAOMOVFIN_OnClick
	End If
End Sub
