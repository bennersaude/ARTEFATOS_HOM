'HASH: 88DFBA40B16A33C1D53F526B05DF7B35
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  Dim SQL As Object
  Dim pListaBenefCanc As String

  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
    	                            0, _
                                   	0, _
                                   	RecordHandleOfTable("SAM_CONTRATO"), _
                                   	vDataFinalSuspensao)Then
   		bsShowMessage("Não é permitido cancelar o contrato por motivo de suspensão!", "E")
   		CanContinue = False
   		Exit Sub
  End If
  Set BSBen001Dll = Nothing


  Set SQL = NewQuery

  SQL.Add("SELECT DATACANCELAMENTO")
  SQL.Add("  FROM SAM_CONTRATO    ")
  SQL.Add(" WHERE HANDLE = :HANDLE")

  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_CONTRATO")


  SQL.Active = True



  Dim Obj As Object
  Set Obj = CreateBennerObject("SAMCANCELAMENTO.Cancelar")

  If VisibleMode Then
	bsShowMessage(Obj.Contrato(CurrentSystem,RecordHandleOfTable("SAM_CONTRATO"),CurrentQuery.FieldByName("DATA").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger, CurrentQuery.FieldByName("DATAATEND").AsDateTime), "I")
  ElseIf WebMode Then
    If SQL.FieldByName("DATACANCELAMENTO").IsNull Then
		bsShowMessage(Obj.Contrato(CurrentSystem, RecordHandleOfTable("SAM_CONTRATO") ,CurrentQuery.FieldByName("DATA").AsDateTime, CurrentQuery.FieldByName("MOTIVO").AsInteger, CurrentQuery.FieldByName("DATAATEND").AsDateTime),  "I")
    Else
 		bsShowMessage("Contrato já cancelado!","I")
 	End If
  End If





  Set Obj = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOMOVFIN" Then
		BOTAOMOVFIN_OnClick
	End If
End Sub

Public Sub TABLE_AfterScroll()
	Dim Nulo As String

	If (StrPos("ORACLE", SQLServer) > 0) Then
		Nulo = "NVL"
	ElseIf (StrPos("DB2", SQLServer) > 0) Then
		Nulo = "COALESCE"
	Else
		Nulo = "ISNULL"
	End If
	Dim vCriterio As String
	vCriterio = "HANDLE NOT IN (SELECT " + Nulo + "(MOTIVOFALECIMENTOTITULAR,0) FROM SAM_PARAMETROSBENEFICIARIO " + _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOFALECIMENTO,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOMIGRACAO,0) FROM SAM_PARAMETROSBENEFICIARIO " + _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MIGRACAONREGREG,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOMIGRACAOCORRESPONSAVEL,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(ADAPTACAO,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOCANCELAMENTOINDICADOR,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOFALTADEPENDENTEAGREGADO,0) FROM SAM_PARAMETROSBENEFICIARIO "+ _
                              "  UNION " + _
                              " SELECT " + Nulo + "(MOTIVOCANCMIGRATIVOINATIVO,0) FROM SAM_PARAMETROSBENEFICIARIO)" _

	If VisibleMode Then
		MOTIVO.LocalWhere = vCriterio
	ElseIf WebMode Then
		MOTIVO.WebLocalWhere = vCriterio
	End If
End Sub
