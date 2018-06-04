'HASH: 8F902A3CBD968C79B8DC1530DC318945
'Macro: SAM_ALERTAPRESTADOR
'#Uses "*bsShowMessage"

Dim vFiltro As String
Dim vFiltroFilial As String
Dim vgDataFinal As Date

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "I")
		CanContinue = False
		Exit Sub
	End If

	If CurrentQuery.State = 3 Then
		bsShowMessage("O registro não pode estar em edição", "I")
		Exit Sub
	End If

	Dim sql As Object
	Set sql = NewQuery

	If Not InTransaction Then StartTransaction

	sql.Add("UPDATE SAM_ALERTAPRESTADOR SET USUARIO=:USUARIO, DATA=:DATA WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)

	sql.ParamByName("USUARIO").Value = CurrentUser
	sql.ParamByName("DATA").Value = ServerNow

	sql.ExecSQL

	If InTransaction Then Commit

	CurrentQuery.Active = False
	CurrentQuery.Active = True

	Set sql = Nothing
End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
	'#Uses "*ProcuraPrestador"..ESTAVA COMENTADO
	If CurrentQuery.State = 1 Then
		TABLE_BeforeEdit(ShowPopup)
		If ShowPopup = False Then
			Exit Sub
		End If
	End If

	Dim ProcuraDLL As Variant
	Dim vColunas As String
	Dim vCampos As String
	Dim vCriterio As String
	Dim vHandle As Long
	Dim vUsuario As String

	vUsuario = Str(CurrentUser)

	Set ProcuraDLL = CreateBennerObject("PROCURA.PROCURAR")

	vColunas = "SAM_PRESTADOR.PRESTADOR|SAM_PRESTADOR.Z_NOME|SAM_PRESTADOR.DATACREDENCIAMENTO"
	vColunas = vColunas + "|SAM_CATEGORIA_PRESTADOR.DESCRICAO|ESTADOS.NOME|MUNICIPIOS.NOME"
	vCriterio = "MUNICIPIOPAGAMENTO IN " + vFiltro + vFiltroFilial

	If (VisibleMode And NodeInternalCode = 2020) Or _
		 (WebMode And WebVisionCode = "V_SAM_ALERTAPRESTADOR_499") Then
		vCriterio = vCriterio + " AND ASSOCIACAO = 'S'"
	End If

	vCampos = "CPF/CNPJ|Nome do Prestador|Data Cred.|Categoria|Estados|Município"
	vHandle = ProcuraDLL.Exec(CurrentSystem, "SAM_PRESTADOR|SAM_CATEGORIA_PRESTADOR[SAM_CATEGORIA_PRESTADOR.HANDLE = SAM_PRESTADOR.CATEGORIA]|ESTADOS[ESTADOS.HANDLE = SAM_PRESTADOR.ESTADOPAGAMENTO]|MUNICIPIOS[MUNICIPIOS.HANDLE = SAM_PRESTADOR.MUNICIPIOPAGAMENTO]", vColunas, 2, vCampos, vCriterio, "Prestadores", True, PRESTADOR.Text)
	ShowPopup = False

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
	End If

	ShowPopup = False

	Set ProcuraDLL = Nothing
End Sub

Public Sub TABLE_AfterPost()
	TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
	BOTAOGERAREVENTOS.Visible=False 
	PRESTADOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	DATAINICIAL.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	DATAFINAL.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	DESCRICAO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	AUTORIZACAOACAO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	AUTORIZACAOEXECUTOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	AUTORIZACAOSOLICITANTE.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	MOTIVONEGACAO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	ACAOPAGAMENTO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	PAGAMENTOEXECUTOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	PAGAMENTORECEBEDOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	MOTIVOGLOSA.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	VALIDOPARAMEMBROS.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	AUTORIZACAOLOCALEXEC.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	AUTORIZACAORECEBEDOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	PAGAMENTOLOCALEXEC.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	ALERTATEXTO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	GERAAUDITORIA.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
	CLASSEPRESTADOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull

	If WebMode And _
	   RecordHandleOfTable("SAM_PRESTADOR") > 0 Then
      PRESTADOR.ReadOnly = True
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
		CanContinue = False
		bsShowMessage("Operação cancelada. Usuário diferente", "E")
		Exit Sub
	End If

	'***************** SMS **********************************************************
	Dim Q As Object
	Set Q = NewQuery

	Q.Add("DELETE FROM SAM_ALERTAPRESTADOR_EVENTO WHERE ALERTAPRESTADOR = :HANDLE")

	Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	Q.ExecSQL
	'**************** Fim ***********************************************************
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	vgDataFinal = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

	Dim Msg As String

	vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)

	If vFiltro = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	'se estiver abaixo da carga de filiais filtra os estados daquela filial +controle de acesso
	If RecordHandleOfTable("FILIAIS")>0 Then
		vFilial = Str(RecordHandleOfTable("FILIAIS"))
		vFiltroFilial = " AND SAM_PRESTADOR.HANDLE IN (SELECT HANDLE FROM SAM_PRESTADOR WHERE FILIALPADRAO = " + vFilial + ") "
	Else
		vFiltroFilial = ""
	End If

	On Error GoTo Erro

	If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
		CanContinue = False
		bsShowMessage("Operação cancelada. Usuário diferente", "E")
	End If

	Erro :

	If WebMode Then
		If WebVisionCode = "V_SAM_ALERTAPRESTADOR" Then
			PRESTADOR.ReadOnly = True
		End If
	End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	vFiltro = checkPermissaoFilial(CurrentSystem, "I", "P", Msg)

	If vFiltro = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	'se estiver abaixo da carga de filiais filtra os estados daquela filial +controle de acesso
	If RecordHandleOfTable("FILIAIS")>0 Then
		vFilial = Str(RecordHandleOfTable("FILIAIS"))
		vFiltroFilial = " AND SAM_PRESTADOR.HANDLE IN (SELECT HANDLE FROM SAM_PRESTADOR WHERE FILIALPADRAO = " + vFilial + ") "
	Else
		vFiltroFilial = ""
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vDllRTF2TXT As Object

	If (Not CurrentQuery.FieldByName("DATAFINAL").IsNull)And _
		 (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)Then
		CanContinue = False
		bsShowMessage("A Data final, se informada, deve ser maior ou igual a inicial", "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "N" And _
		 CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "N" Then
		CanContinue = False
		bsShowMessage("Pelo menos uma ação deve ser diferente de Nada", "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString <>"N" And _
		 CurrentQuery.FieldByName("AUTORIZACAOEXECUTOR").AsString = "N" And _
		 CurrentQuery.FieldByName("AUTORIZACAOSOLICITANTE").AsString = "N" And _
		 CurrentQuery.FieldByName("AUTORIZACAOLOCALEXEC").AsString = "N" And _
		 CurrentQuery.FieldByName("AUTORIZACAORECEBEDOR").AsString = "N" Then
		CanContinue = False
		bsShowMessage("Executor, Solicitante, Executor ou e/ou Local Execução para autorização deve ser selecionado", "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString <>"N" And _
		 CurrentQuery.FieldByName("PAGAMENTOEXECUTOR").AsString = "N" And _
		 CurrentQuery.FieldByName("PAGAMENTOLOCALEXEC").AsString = "N" And _
		 CurrentQuery.FieldByName("PAGAMENTORECEBEDOR").AsString = "N" Then
		CanContinue = False
		bsShowMessage("Executor, Recebedor ou Local Execução para pagamento deve ser selecionado", "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "R" And _
		 CurrentQuery.FieldByName("MOTIVONEGACAO").IsNull Then
		CanContinue = False
		bsShowMessage("Para alerta de restrição na autorização deve ser informado o motivo de negação", "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "R" And _
		 CurrentQuery.FieldByName("MOTIVOGLOSA").IsNull Then
		CanContinue = False
		bsShowMessage("Para alerta de restrição no pagamento deve ser informado o motivo de glosa", "E")
		Exit Sub
	End If

	Set vDllRTF2TXT = CreateBennerObject("RTF2TXT.Rotinas")

	CurrentQuery.FieldByName("ALERTATEXTOTXT").AsString = vDllRTF2TXT.Rtf2Txt(CurrentSystem, CurrentQuery.FieldByName("ALERTATEXTO").AsString)

	'SMS 59169 - Marcelo Barbosa - 17/03/2006
	'If InStr(CurrentQuery.FieldByName("ALERTATEXTOTXT").AsString,"{") > 0 Or _
	'	 InStr(CurrentQuery.FieldByName("ALERTATEXTOTXT").AsString,"}") > 0 Then
	'	bsShowMessage("Não é permitido inserir os caracteres { (abre chave) e/ou } (fecha chave) no texto do Alerta!", "E")
	'	CanContinue = False
	'	Exit Sub
	'End If
	'Fim - SMS 59169

	Set vDllRTF2TXT = Nothing

    If vgDataFinal <>CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
      If VisibleMode Then
        If bsShowMessage("Fechando a vigência não será permitido alteração no alerta , nem reabrir a vigência." + (Chr(13)) + _
                         "Deseja continuar?", "Q") = vbNo Then
          CanContinue = False
          Exit Sub
        End If
      Else
        bsShowMessage("A vigência foi fechada. Não será permitida a alteração do alerta!", "I")
      End If
    End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOALTERARRESPONSAVEL"
			BOTAOALTERARRESPONSAVEL_OnClick
	End Select
End Sub
