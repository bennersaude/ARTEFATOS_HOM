'HASH: 567E8AEE47A976E8A5845B67E2EF15F8
'Macro: SAM_PLANO_PFPRESTADOR
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub BOTAOATUALIZACONTRATO_OnClick()
	Dim q1 As Object
	Set q1 = NewQuery
	Dim q2 As Object
	Set q2 = NewQuery
	Dim q3 As Object
	Set q3 = NewQuery
	Dim q4 As Object '25/04/2003 -busca os contratos que já tenham o plano em questão cadastrado -Roger
	Set q4 = NewQuery
	Dim q5 As Object '25/04/2003 -atualiza a situação desses mesmos contratos,com relação ao plano em questão -Roger
	Set q5 = NewQuery
	Dim HandleFiltro As Long
	Dim Filtro As Object
	Set Filtro = CreateBennerObject("SamFiltro.Filtro")

	Dim DataInicial As Date
	Dim DataFinal As Date

	inicio :
		If (WebMode) Then
			If CurrentVirtualQuery.FieldCount > 0 Then
				DataInicial = CurrentVirtualQuery.FieldByName("DATAINICIAL").AsDateTime
				DataFinal = CurrentVirtualQuery.FieldByName("DATAFINAL").AsDateTime
			End If

			If(DataInicial > DataFinal)Then
				bsShowMessage("A Data final não pode ser anterior à data inicial!", "I")
				CanContinue = False
				Exit Sub
			End If
		Else
			HandleFiltro = Filtro.Exec(CurrentSystem, CurrentUser, 801, "DATAINICIAL.ob|DATAFINAL", "Copiar PF")

			If (HandleFiltro <= 0) Then
				Set q1 = Nothing
				Set q2 = Nothing
				Set q3 = Nothing
				Set q4 = Nothing
				Set q5 = Nothing
				Set Filtro = Nothing
				Exit Sub
			End If

			q1.Add("SELECT DATAINICIAL, DATAFINAL    ")
			q1.Add("FROM RF_FILTRO                   ")
			q1.Add("WHERE HANDLE=:PHANDLEFILTRO                  ")

			q1.ParamByName("PHANDLEFILTRO").Value = HandleFiltro
				q1.Active = True

			If (Not q1.FieldByName("DATAFINAL").IsNull) And _
		   	(q1.FieldByName("DATAFINAL").AsDateTime < q1.FieldByName("DATAINICIAL").AsDateTime) Then
				bsShowMessage("Data final não pode ser anterior à data inicial.", "I")
				GoTo inicio
			End If

			DataInicial = q1.FieldByName("DATAINICIAL").AsDateTime
			DataFinal = q1.FieldByName("DATAFINAL").AsDateTime
		End If
		q2.Clear
		q2.Add("SELECT HANDLE CONTRATO FROM SAM_CONTRATO C")
		q2.Add("WHERE C.DATACANCELAMENTO IS NULL")
		q2.Add("AND C.PLANO=:pPLANO")
		q2.Add("AND HANDLE NOT IN ")
		q2.Add("	(SELECT CONTRATO FROM SAM_CONTRATO_PFPRESTADOR")
		q2.Add("	 WHERE PRESTADOR=:pPRESTADOR)")
		q2.ParamByName("pPLANO").Value = CurrentQuery.FieldByName("PLANO").Value
		q2.ParamByName("pPRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
		q2.Active = True

		q4.Clear
		q4.Add("SELECT HANDLE CONTRATO")
		q4.Add("  FROM SAM_CONTRATO_PFPRESTADOR")
		q4.Add(" WHERE PRESTADOR=:pPRESTADOR")
		q4.ParamByName("pPRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").Value
		q4.Active = True

		If Not InTransaction Then
			StartTransaction
		End If

		q3.Clear
		q3.Add("INSERT INTO SAM_CONTRATO_PFPRESTADOR																")
		q3.Add("(HANDLE, PRESTADOR, EVENTO, CONTRATO, OBSERVACAO, CODIGOPF, ACEITAPARCELAMENTO, GRAU, DATAINICIAL, DATAFINAL)	")
		q3.Add("VALUES																							")
		q3.Add("(:pHANDLE, :pPRESTADOR, :pEVENTO, :pCONTRATO, :pOBSERVACAO, :pCODIGOPF, :pACEITAPARCELAMENTO, :pGRAU, :pDATAINICIAL, :pDATAFINAL)")

		On Error GoTo FIM

		While Not q2.EOF
			q3.ParamByName("pHANDLE").AsInteger = NewHandle("SAM_CONTRATO_PFPRESTADOR")
			q3.ParamByName("pPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
			q3.ParamByName("pEVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
			q3.ParamByName("pCONTRATO").AsInteger = q2.FieldByName("CONTRATO").AsInteger
			q3.ParamByName("pOBSERVACAO").AsMemo = CurrentQuery.FieldByName("OBSERVACAO").Value
			q3.ParamByName("pCODIGOPF").AsInteger = CurrentQuery.FieldByName("CODIGOPF").AsInteger
			q3.ParamByName("pACEITAPARCELAMENTO").AsString = CurrentQuery.FieldByName("ACEITAPARCELAMENTO").AsString
			q3.ParamByName("pGRAU").DataType = ftInteger
			q3.ParamByName("pGRAU").Value = IIf(CurrentQuery.FieldByName("GRAU").IsNull, Null, CurrentQuery.FieldByName("GRAU").AsInteger)
			q3.ParamByName("pDATAINICIAL").AsDateTime = DataInicial

			If IsNull(DataFinal) Then
				q3.ParamByName("pDATAFINAL").DataType = ftDateTime
				q3.ParamByName("pDATAFINAL").Clear
			Else
				q3.ParamByName("pDATAFINAL").AsDateTime = DataFinal
			End If

			q3.ExecSQL
			q2.Next
		Wend
		While Not q4.EOF
			q5.Clear
			q5.Add("UPDATE SAM_CONTRATO_PFPRESTADOR SET")
			q5.Add("EVENTO=:pEVENTO1,")
			q5.Add("OBSERVACAO=:pOBSERVACAO1,")
			q5.Add("CODIGOPF=:pCODIGOPF1,")
			q5.Add("ACEITAPARCELAMENTO=:pACEITAPARCELAMENTO1,")
			q5.Add("GRAU=:pGRAU1,")
			q5.Add("DATAINICIAL=:pDATAINICIAL1,")
			q5.Add("DATAFINAL=:pDATAFINAL1")
			q5.Add("WHERE HANDLE=:pHANDLE1")
			q5.ParamByName("pHANDLE1").AsInteger = q4.FieldByName("CONTRATO").AsInteger
			q5.ParamByName("pEVENTO1").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
			q5.ParamByName("pOBSERVACAO1").AsMemo = CurrentQuery.FieldByName("OBSERVACAO").Value
			q5.ParamByName("pCODIGOPF1").AsInteger = CurrentQuery.FieldByName("CODIGOPF").AsInteger
			q5.ParamByName("pACEITAPARCELAMENTO1").AsString = CurrentQuery.FieldByName("ACEITAPARCELAMENTO").AsString
			q5.ParamByName("pGRAU1").DataType = ftInteger
			q5.ParamByName("pGRAU1").Value = IIf(CurrentQuery.FieldByName("GRAU").IsNull, Null, CurrentQuery.FieldByName("GRAU").AsInteger)
			q5.ParamByName("pDATAINICIAL1").AsDateTime = DataIncial
			If IsNull(DataFinal) Then
				q5.ParamByName("pDATAFINAL1").DataType = ftDateTime
				q5.ParamByName("pDATAFINAL1").Clear
			Else
				q5.ParamByName("pDATAFINAL1").AsDateTime = DataFinal
			End If
				q5.ExecSQL
				q4.Next
		Wend
		Commit

		Set q1 = Nothing
		Set q2 = Nothing
		Set q3 = Nothing
		Set q4 = Nothing
		Set q5 = Nothing
		bsShowMessage("Atualização completa ", "I")
		Exit Sub

	FIM :
		Rollback
		bsShowMessage("Ocorreu o seguinte erro ao atualizar contrato(s)" + Str(Error), "I")

		Set q1 = Nothing
		Set q2 = Nothing
		Set q3 = Nothing
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
	'#Uses "*ProcuraPrestador"
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraPrestador("C", "T", PRESTADOR.Text)' pelo CPF e Todos

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
	End If
End Sub
Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T3880" Then
			PRESTADOR.ReadOnly = True
		End If
		If WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If
		If WebMenuCode = "T1612" Then
			PLANO.ReadOnly = True
		End If
	End If

	BOTAOATUALIZACONTRATO.Enabled = IIf(CurrentQuery.State = 1, True, False)

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOATUALIZACONTRATO"
			BOTAOATUALIZACONTRATO_OnClick
	End Select
End Sub
