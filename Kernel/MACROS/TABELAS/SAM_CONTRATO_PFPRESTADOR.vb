'HASH: 5BF0835E9FE99854D5FACAB262F75ECD
'Macro: SAM_CONTRATO_PFPRESTADOR
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"

Option Explicit

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
	Dim interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String

	ShowPopup = False

	Set interface = CreateBennerObject("Procura.Procurar")

	vColunas = "CONTRATO|CONTRATANTE|SAM_PLANO.DESCRICAO"
	vCriterio = ""
	vCampos = "Contrato|Contratante|Plano"
	vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO|SAM_PLANO[SAM_PLANO.HANDLE=SAM_CONTRATO.PLANO]", vColunas, 1, vCampos, vCriterio, "Contratos", True, "")

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CONTRATO").Value = vHandle
	End If

	Set interface = Nothing
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

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraGrau(GRAU.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
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

Public Sub TABLE_AfterPost()
	TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
	If VisibleMode Then
		PLANO.LocalWhere = "SAM_PLANO.HANDLE IN (SELECT PLANO " + _
						   "					   FROM SAM_CONTRATO_PLANO " + _
						   "					  WHERE CONTRATO = @CONTRATO)"
	Else
		PLANO.WebLocalWhere = "A.HANDLE IN (SELECT PLANO " + _
							  "				  FROM SAM_CONTRATO_PLANO " + _
							  "				 WHERE CONTRATO = @CAMPO(CONTRATO))"
	End If

	If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		DATAFINAL.ReadOnly = False
	Else
		DATAFINAL.ReadOnly = True
	End If

	If WebMode Then
		If WebMenuCode = "T2886" Then
			CONTRATO.ReadOnly = True
		End If
		If WebMenuCode = "T3880" Then
			PRESTADOR.ReadOnly = True
		End If
		If WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If
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

	qContrato.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
	qContrato.Active = True

	'Não permitir a inclusão de registros com vigência iniciando antes da adesão do contrato ou terminando depois do cancelamento do mesmo (se houver).
	If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < qContrato.FieldByName("DATAADESAO").AsDateTime) Then
		bsShowMessage("A data inicial não pode ser anterior à adesão do contrato.","E")
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
	Dim interface As Object
	Dim Linha As String
	Dim Condicao As String

	'Balani SMS 4595 18/08/2005
	Condicao = "AND PRESTADOR = " + CurrentQuery.FieldByName("PRESTADOR").AsString + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString 'Anderson sms 21638
	Condicao = Condicao + " AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString

	If CurrentQuery.FieldByName("REGRA").AsInteger > 0 Then
		Condicao = Condicao + " AND REGRA = " + CurrentQuery.FieldByName("REGRA").AsString
	Else
		Condicao = Condicao + " AND REGRA IS NULL"
	End If
	'final Balani SMS 4595

	Set interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = interface.Vigencia(CurrentSystem, "SAM_CONTRATO_PFPRESTADOR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATO", Condicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If
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

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		CanContinue = False
		bsShowMessage("Registro finalizado não pode ser alterado!", "E")
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
