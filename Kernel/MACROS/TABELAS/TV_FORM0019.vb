'HASH: 652A84AF1749D15FD6A2976864FCCFDB
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOMOVFIN_OnClick()
	Dim qAux As Object

	Set qAux = NewQuery

	qAux.Add("SELECT HANDLE FROM R_RELATORIOS")
	qAux.Add(" WHERE CODIGO = 'BEN070'")
 	qAux.Active = True

	ReportPreview(qAux.FieldByName("HANDLE").AsInteger, "", False, False)

 	Set qAux = Nothing
End Sub


Public Sub MOTIVO_OnChange()
	If Not CurrentQuery.FieldByName("MOTIVO").IsNull Then
		Dim qParamBenef As Object

		Set qParamBenef = NewQuery

		qParamBenef.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")

		qParamBenef.Active = True

		If CurrentQuery.FieldByName("MOTIVO").AsInteger = qParamBenef.FieldByName("MOTIVOFALECIMENTO").AsInteger Then
			CID.Visible = True
			DATAFALECIMENTO.Visible = True
		Else
			CID.Visible = False
			DATAFALECIMENTO.Visible = False
		End If

		Set qParamBenef = Nothing
	End If
End Sub

Public Sub TABLE_AfterScroll()
	Dim Nulo As String

	If VisibleMode Then
		CID.Visible = False
		DATAFALECIMENTO.Visible = False
		CHECKEMITIRCOMUNICADO.Visible = False
	End If

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

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  SessionVar("CANCBENEF_PARAMESPECIFICO") = ""
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qParamBenef As Object
	Set qParamBenef = NewQuery

	UserVar("LISTABENEFCANC") = ""


	qParamBenef.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")
	qParamBenef.Active = True

	If CurrentQuery.FieldByName("MOTIVO").AsInteger = qParamBenef.FieldByName("MOTIVOFALECIMENTO").AsInteger Then
		If CurrentQuery.FieldByName("DATAFALECIMENTO").IsNull Then
			bsShowMessage("A Data de Falecimento deve ser informada!", "E")
			CanContinue = False
			Exit Sub
		End If
	End If

	Dim vcContainer As CSDContainer
	Dim vbEspec As Boolean

	If SessionVar("CANCBENEF_PARAMESPECIFICO") <> "" Then
		Set vcContainer = NewContainer

		vcContainer.SetXML(SessionVar("CANCBENEF_PARAMESPECIFICO"), True, False, True)

		vbEspec = True
	Else
		vbEspec = False
	End If

	If Not CurrentQuery.FieldByName("DATAATEND").IsNull Then
		If CurrentQuery.FieldByName("DATAATEND").AsDateTime <= CurrentQuery.FieldByName("DATA").AsDateTime Then
			bsShowMessage("A data de atendimento até deve ser posterior a data de cancelamento.","E")

			CanContinue = False
			Set vcContainer = Nothing

			Exit Sub
		End If
	End If

    Dim dllBSBen001 As Object

    Set dllBSBen001 = CreateBennerObject("BSBen001.Beneficiario")

    Dim vsMensagem  As String

    If dllBSBen001.ValidarCancelamentoManual(CurrentSystem, _
                                             RecordHandleOfTable("SAM_BENEFICIARIO"), _
                                             vsMensagem) Then
      bsShowMessage(vsMensagem, "E")
      CanContinue = False
    Else
	  Dim Obj                            As Object
	  Dim vbCancelar					 As Boolean
	  Set Obj = CreateBennerObject("SAMCANCELAMENTO.Cancelar")

	  If Obj.ChecaInternado(CurrentSystem, CLng(SessionVar("HBENEFICIARIO")), CurrentQuery.FieldByName("DATA").AsDateTime) = True Then

		If bsShowMessage("Beneficiário titular com autorização de internação em aberto. Deseja realmente cancelar?", "Q") = vbYes Then
		  vbCancelar = True
		Else
		  vbCancelar = False
		  bsShowMessage("Beneficiário titular com autorização de internação em aberto. Cancelamento abortado.", "E")
		End If
	  Else
	  	vbCancelar = True
	  End If

	  If vbCancelar = True Then
    	If vbEspec Then
		  bsShowMessage(Obj.Beneficiario(CurrentSystem, _
		    						     CLng(SessionVar("HBENEFICIARIO")), _
									     CurrentQuery.FieldByName("DATA").AsDateTime, _
									     CurrentQuery.FieldByName("MOTIVO").AsInteger, _
									     CurrentQuery.FieldByName("DATAATEND").AsDateTime, _
									     0, _
									     True, _
									     CurrentQuery.FieldByName("CHECKCANCDOCFAT").AsBoolean, _
									     CurrentQuery.FieldByName("CID").AsInteger, _
									     CurrentQuery.FieldByName ("DATAFALECIMENTO").AsDateTime, _
									     vcContainer), _
						                 "I")
	    Else
		  bsShowMessage(Obj.Beneficiario(CurrentSystem, _
									     CLng(SessionVar("HBENEFICIARIO")), _
									     CurrentQuery.FieldByName("DATA").AsDateTime, _
									     CurrentQuery.FieldByName("MOTIVO").AsInteger, _
									     CurrentQuery.FieldByName("DATAATEND").AsDateTime, _
									     0, _
									     True, _
									     CurrentQuery.FieldByName("CHECKCANCDOCFAT").AsBoolean, _
									     CurrentQuery.FieldByName("CID").AsInteger, _
									     CurrentQuery.FieldByName ("DATAFALECIMENTO").AsDateTime, _
									     Null), _
						                 "I")
	    End If
	  End If
	End If

	If WebMode And UserVar("LISTABENEFCANC") <> "" Then
		If(CurrentQuery.FieldByName("CHECKEMITIRCOMUNICADO").AsBoolean) Then
			EMITIRCOMUNICADO
		End If
	End If

	Set Obj         = Nothing
	Set dllBSBen001 = Nothing
	Set vcContainer = Nothing
End Sub

Public Sub EMITIRCOMUNICADO()
	Dim Obj As Object
    Set Obj = CreateBennerObject("SAMCANCELAMENTO.Cancelar")
	Dim qRelatorio As Object
	Set qRelatorio = NewQuery
	Dim HandleRelatorio As Integer


	Dim qBeneficiario As Object
	Set qBeneficiario = NewQuery
	qBeneficiario.Add("SELECT CONTRATO FROM SAM_BENEFICIARIO WHERE HANDLE = "+Str( RecordHandleOfTable("SAM_BENEFICIARIO") ) )
	qBeneficiario.Active = True

	Dim qContrato As Object
	Set qContrato = NewQuery
	qContrato.Add("SELECT * FROM SAM_CONTRATO WHERE HANDLE = "+Str( qBeneficiario.FieldByName("CONTRATO").AsInteger) )
	qContrato.Active = True

	Dim qRelConvenio As Object
	Set qRelConvenio = NewQuery
	qRelConvenio.Add("SELECT RELATORIOAVISOCANCELAMENTO FROM SAM_CONVENIO WHERE HANDLE = "+Str(qContrato.FieldByName("CONVENIO").AsInteger))
	qRelConvenio.Active = True


	Dim CodRelatorio As String

	If qRelConvenio.FieldByName("RELATORIOAVISOCANCELAMENTO").AsString = "" Then
		CodRelatorio = "'WEB.BEN003B'"
	Else
		CodRelatorio = "'WEB."+qRelConvenio.FieldByName("RELATORIOAVISOCANCELAMENTO").AsString+"'"
	End If

	qRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = " + CodRelatorio)
	qRelatorio.Active = True
	HandleRelatorio = qRelatorio.FieldByName("HANDLE").AsInteger

	Dim relatorio As CSReportPrinter
	Set relatorio = NewReport(HandleRelatorio)
	relatorio.SqlWhere = "A.HANDLE IN " + UserVar("LISTABENEFCANC")
	relatorio.Preview


	UserVar("LISTABENEFCANC") = ""
	Set qContrato = Nothing
	Set qRelConvenio = Nothing
	Set qRelatorio = Nothing
	Set qBeneficiario = Nothing
	Set relatorio = Nothing
End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOMOVFIN"
			BOTAOMOVFIN_OnClick
		Case "CMDEMITIRCOMUNICADO"
			EMITIRCOMUNICADO
	End Select
End Sub
