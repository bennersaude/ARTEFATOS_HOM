'HASH: 67D467E76CD70D92825C4F8AD4840565
'#Uses "*bsShowMessage"
'#Uses "*CheckCPFCNPJ"



Public Sub BOTAORECUSARREGISTRO_OnClick()

	Dim vsSqlParamGerais As Object
	Dim vsSqlUpdate As Object

	Set vsSqlParamGerais = NewQuery
	Set vsSqlUpdate = NewQuery

	vsSqlParamGerais.Clear
	vsSqlParamGerais.Active = False
	vsSqlParamGerais.Add("SELECT PRAZOENVIODOCUMENTOS FROM SAM_PARAMETROSWEB")
    vsSqlParamGerais.Active = True

	vsSqlUpdate.Clear
	vsSqlUpdate.Add("UPDATE WEB_SAM_BENEFICIARIO SET SITUACAO = '3' WHERE (SITUACAO = '1' OR SITUACAO = '5') AND DATAHORAINCLUSAO + :DIAS < :DATAFINAL")
	vsSqlUpdate.ParamByName("DIAS").AsInteger = vsSqlParamGerais.FieldByName("PRAZOENVIODOCUMENTOS").AsInteger
	vsSqlUpdate.ParamByName("DATAFINAL").AsDateTime = ServerNow
	vsSqlUpdate.ExecSQL

  	bsShowMessage("Registros rejeitados com sucesso!", "I")

	RefreshNodesWithTable("WEB_SAM_BENEFICIARIO")

	Set vsSqlParamGerais = Nothing
	Set vsSqlUpdate      = Nothing

End Sub

Public Sub OPERACAODESKTOP_OnChange()

	Dim interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabelas As String

	ShowPopup = False
	Set interface = CreateBennerObject("Procura.Procurar")


	LimparCampos
	CurrentQuery.UpdateRecord

	If CurrentQuery.FieldByName("OPERACAODESKTOP").Value = "A" Then

		vColunas = "NOME|BENEFICIARIO"

		vCriterio = "DATACANCELAMENTO IS NULL"
		vCampos = "Nome|Beneficiário"

		vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 1, vCampos, vCriterio, "Beneficiário", False, "")

		If vHandle <> 0 Then

			PopularCamposAPartirDeBeneficiarioSelecionado(vHandle)
		Else
			CurrentQuery.FieldByName("OPERACAODESKTOP").Clear
			LimparCampos
		End If


	ElseIf CurrentQuery.FieldByName("OPERACAODESKTOP").Value = "I" Then
		vColunas = "SAM_FAMILIA.FAMILIA|SAM_CONTRATO.CONTRATO|SAM_CONTRATO.CONTRATANTE"

		vCriterio = "SAM_FAMILIA.DATACANCELAMENTO IS NULL"

		vCampos = "Família|Contrato|Contratante"
		vTabelas = "SAM_FAMILIA|SAM_CONTRATO[SAM_FAMILIA.CONTRATO = SAM_CONTRATO.HANDLE]"
		vHandle = interface.Exec(CurrentSystem, vTabelas, vColunas, 1, vCampos, vCriterio, "Família", False, "")

		If vHandle <> 0 Then

			PopularCamposParaInclusao(vHandle)

		Else
			LimparCampos
		End If


	End If

	Set interface = Nothing

	ParentescoDeveFicarReadOnly

End Sub

Public Sub LimparCampos

	CurrentQuery.FieldByName("NOME").Clear
	CurrentQuery.FieldByName("CPF").Clear
	CurrentQuery.FieldByName("DATANASCIMENTO").Clear
	CurrentQuery.FieldByName("MATRICULAFUNCIONAL").Clear
	CurrentQuery.FieldByName("ESTADOCIVIL").Clear
	CurrentQuery.FieldByName("SEXO").Clear
	CurrentQuery.FieldByName("PARENTESCO").Clear
	CurrentQuery.FieldByName("TIPOBENEFICIARIO").Clear
	CurrentQuery.FieldByName("DATAMATRIMONIO").Clear
	CurrentQuery.FieldByName("CARTAONACIONALSAUDE").Clear
	CurrentQuery.FieldByName("DNV").Clear
	CurrentQuery.FieldByName("NOMEMAE").Clear
	CurrentQuery.FieldByName("PISPASEP").Clear
	CurrentQuery.FieldByName("ATIVIDADEPRINCIPAL").Clear
	CurrentQuery.FieldByName("CONTRATO").Clear
	CurrentQuery.FieldByName("DATAADMISSAO").Clear
	CurrentQuery.FieldByName("FAMILIA").Clear
	CurrentQuery.FieldByName("DOCUMENTO").Clear
	CurrentQuery.FieldByName("TIPODOCUMENTO").Clear
	CurrentQuery.FieldByName("DATAEMISSAO").Clear
	CurrentQuery.FieldByName("ORGAOEMISSOR").Clear
End Sub


Public Sub PARENTESCO_OnPopup(ShowPopup As Boolean)

	If Not CurrentQuery.FieldByName("OPERACAODESKTOP").IsNull Then

		Dim interface As Object
		Dim vHandle As Long
		Dim vCampos As String
		Dim vColunas As String
		Dim vCriterio As String
		Dim vTabelas As String

		ShowPopup = False
		Set interface = CreateBennerObject("Procura.Procurar")
		vColunas = "CODIGO|DESCRICAO"

		vCriterio = "  HANDLE IN (SELECT TIPODEPENDENTE "
		vCriterio = vCriterio  + "  FROM SAM_CONTRATO_TPDEP "
		vCriterio = vCriterio  + " WHERE CONTRATO = " + CurrentQuery.FieldByName("CONTRATO").AsString
		vCriterio = vCriterio  + "   AND DATAINICIAL <= GETDATE() AND ((DATAFINAL Is NULL) Or (DATAFINAL >= GETDATE() )))"

		vCriterio = vCriterio  + " AND GRUPODEPENDENTE <> 'T' "




		vCampos = "Código|Descrição"

		vHandle = interface.Exec(CurrentSystem, "SAM_TIPODEPENDENTE", vColunas, 1, vCampos, vCriterio, "Beneficiário", True, "")

		If vHandle > 0 Then
			CurrentQuery.FieldByName("PARENTESCO").AsInteger = vHandle
		End If

		Set interface = Nothing

	Else
		MsgBox("Selecione a operação.")
		ShowPopup = False
		Exit Sub
	End If
	NOME.SetFocus


End Sub

Public Sub TABLE_AfterInsert()

	CurrentQuery.FieldByName("MODULO").AsString = "N"
	CurrentQuery.FieldByName("USUARIOINCLUSAO").AsInteger = CurrentUser
	CurrentQuery.FieldByName("DATAHORAINCLUSAO").AsDateTime = ServerDate
	CurrentQuery.FieldByName("SITUACAO").AsInteger = 1

End Sub

Public Sub TABLE_AfterPost()
	Dim bs As CSBusinessComponent

	If CurrentQuery.FieldByName("OPERACAODESKTOP").Value = "A" Then

		Set bs = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.WEB.WebSamBeneficiarioBLL, Benner.Saude.Beneficiarios.Business") ' formato: [namespace.classe], [assembly]
		bs.ClearParameters
		bs.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
		CStr(bs.Execute("SalvarAlteracoesCadastrais"))
		Set bs = Nothing
	ElseIf CurrentQuery.FieldByName("OPERACAODESKTOP").Value = "I" Then

		Set bs = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.WEB.WebSamBeneficiarioBLL, Benner.Saude.Beneficiarios.Business") ' formato: [namespace.classe], [assembly]
		bs.ClearParameters
		bs.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
		CStr(bs.Execute("IncluirBeneficiario"))
		Set bs = Nothing

	End If

End Sub

Public Sub TABLE_AfterScroll()
	' Permite alterar Operação apenas quando for inclusão.
	If CurrentQuery.State = 3 Then
		OPERACAODESKTOP.ReadOnly = False
	Else
		OPERACAODESKTOP.ReadOnly = True

	End If

	If Not CurrentQuery.FieldByName("SITUACAO").IsNull Then
		If CurrentQuery.FieldByName("SITUACAO").AsInteger = 1 Then
			CamposEditaveis
			ParentescoDeveFicarReadOnly
		Else
			CamposSomenteLeitura
		End If

	End If

	SITUACAO.ReadOnly = True
	CONTRATO.ReadOnly = True
	FAMILIA.ReadOnly = True


End Sub

Public Sub CamposSomenteLeitura
   	NOME.ReadOnly = True
	CPF.ReadOnly = True
	DATANASCIMENTO.ReadOnly = True
	MATRICULAFUNCIONAL.ReadOnly = True
	ESTADOCIVIL.ReadOnly = True
	SEXO.ReadOnly = True
	PARENTESCO.ReadOnly = True
	NOMEMAE.ReadOnly = True
	PISPASEP.ReadOnly = True
	DATAADMISSAO.ReadOnly = True
	DOCUMENTO.ReadOnly = True
	TIPODOCUMENTO.ReadOnly = True
	DATAEMISSAO.ReadOnly = True
	PARENTESCO.ReadOnly = True
	ORGAOEMISSOR.ReadOnly = True

End Sub

Public Sub CamposEditaveis
	NOME.ReadOnly = False
	CPF.ReadOnly = False
	DATANASCIMENTO.ReadOnly = False
	MATRICULAFUNCIONAL.ReadOnly = False
	ESTADOCIVIL.ReadOnly = False
	SEXO.ReadOnly = False
	PARENTESCO.ReadOnly = False
	NOMEMAE.ReadOnly = False
	PISPASEP.ReadOnly = False
	DATAADMISSAO.ReadOnly = False
	DOCUMENTO.ReadOnly = False
	TIPODOCUMENTO.ReadOnly = False
	DATAEMISSAO.ReadOnly = False
	ORGAOEMISSOR.ReadOnly = False
End Sub

Public Sub ParentescoDeveFicarReadOnly

	If CurrentQuery.FieldByName("OPERACAODESKTOP").AsString = "A" Then
		Dim qTipoDependente As Object
		Set qTipoDependente = NewQuery
		qTipoDependente.Active = False
		qTipoDependente.Add("SELECT GRUPODEPENDENTE FROM SAM_TIPODEPENDENTE WHERE HANDLE = :HANDLE ")
		qTipoDependente.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PARENTESCO").AsInteger
		qTipoDependente.Active = True

		If qTipoDependente.FieldByName("GRUPODEPENDENTE").AsString = "T" Then
			PARENTESCO.ReadOnly = True
		Else
			PARENTESCO.ReadOnly = False
		End If

		Set qTipoDependente = Nothing

	Else
		PARENTESCO.ReadOnly = False
	End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	CamposEditaveis
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	CurrentQuery.FieldByName("OPERACAO").AsString = CurrentQuery.FieldByName("OPERACAODESKTOP").AsString


	Dim Msg As String
    If Not CurrentQuery.FieldByName("CPF").IsNull Then
    	If Not CheckCPFCNPJ(CurrentQuery.FieldByName("CPF").AsString, 0, True, Msg) Then
    	  	bsShowMessage(Msg, "E")
      		CanContinue = False
    	End If
  	End If

  	If (CurrentQuery.FieldByName("OPERACAODESKTOP").AsString = "I") Then

		If (CurrentQuery.FieldByName("PARENTESCO").IsNull) Then
    	  	bsShowMessage("Para inclusão, o campo parentesco é obrigatório!", "E")
      		CanContinue = False
      		Exit Sub
		End If

  	End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAORECUSARREGISTRO"
		  BOTAORECUSARREGISTRO_OnClick
	End Select

End Sub

Public Sub PopularCamposAPartirDeBeneficiarioSelecionado(pHandleBeneficiario As Long)
	Dim qBeneficiario As Object
	Set qBeneficiario = NewQuery
	qBeneficiario.Active = False
	qBeneficiario.Add("SELECT * FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE ")
	qBeneficiario.ParamByName("HANDLE").AsInteger = pHandleBeneficiario
	qBeneficiario.Active = True


	Dim qMatricula As Object
	Set qMatricula = NewQuery
	qMatricula.Active = False
	qMatricula.Add("SELECT * FROM SAM_MATRICULA WHERE HANDLE = :HANDLE ")
	qMatricula.ParamByName("HANDLE").AsInteger = qBeneficiario.FieldByName("MATRICULA").AsInteger
	qMatricula.Active = True


	CurrentQuery.FieldByName("NOME").AsString = qBeneficiario.FieldByName("NOME").AsString
	CurrentQuery.FieldByName("CPF").AsString = qMatricula.FieldByName("CPF").AsString
	CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime = qMatricula.FieldByName("DATANASCIMENTO").AsDateTime
	CurrentQuery.FieldByName("MATRICULAFUNCIONAL").AsString = qBeneficiario.FieldByName("MATRICULAFUNCIONAL").AsString
	CurrentQuery.FieldByName("ESTADOCIVIL").AsInteger = qBeneficiario.FieldByName("ESTADOCIVIL").AsInteger
	CurrentQuery.FieldByName("SEXO").AsString = qMatricula.FieldByName("SEXO").AsString
	CurrentQuery.FieldByName("PARENTESCO").AsInteger = RetornarTipoParentesco(qBeneficiario.FieldByName("TIPODEPENDENTE").AsInteger)
	CurrentQuery.FieldByName("TIPOBENEFICIARIO").AsString = RetornarTipoBeneficiario(CurrentQuery.FieldByName("PARENTESCO").AsInteger)

	If Not qMatricula.FieldByName("DATACASAMENTO").IsNull Then
		CurrentQuery.FieldByName("DATAMATRIMONIO").AsDateTime = qMatricula.FieldByName("DATACASAMENTO").AsDateTime
	End If

	CurrentQuery.FieldByName("CARTAONACIONALSAUDE").AsString = qMatricula.FieldByName("CARTAONACIONALSAUDE").AsString
	CurrentQuery.FieldByName("DNV").AsString = qMatricula.FieldByName("DNV").AsString
	CurrentQuery.FieldByName("NOMEMAE").AsString = qMatricula.FieldByName("NOMEMAE").AsString
	CurrentQuery.FieldByName("PISPASEP").AsString = qMatricula.FieldByName("PISPASEP").AsString

	If Not qBeneficiario.FieldByName("CBO").IsNull Then
		CurrentQuery.FieldByName("ATIVIDADEPRINCIPAL").AsInteger = qBeneficiario.FieldByName("CBO").AsInteger
	End If

	If Not qBeneficiario.FieldByName("DATAADMISSAO").IsNull Then
		CurrentQuery.FieldByName("DATAADMISSAO").AsDateTime = qBeneficiario.FieldByName("DATAADMISSAO").AsDateTime
	End If

	If Not qMatricula.FieldByName("RG").IsNull Then
		CurrentQuery.FieldByName("DOCUMENTO").AsString = qMatricula.FieldByName("RG").AsString
		CurrentQuery.FieldByName("TIPODOCUMENTO").AsString = "1"

		If Not qMatricula.FieldByName("DATAEXPEDICAORG").IsNull Then
			CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime = qMatricula.FieldByName("DATAEXPEDICAORG").AsDateTime
		End If

		CurrentQuery.FieldByName("ORGAOEMISSOR").AsString = qMatricula.FieldByName("ORGAOEMISSOR").AsString

	ElseIf Not qMatricula.FieldByName("PASSAPORTE").IsNull Then
		CurrentQuery.FieldByName("DOCUMENTO").AsString = qMatricula.FieldByName("PASSAPORTE").AsString
		CurrentQuery.FieldByName("TIPODOCUMENTO").AsString = "2"

		If Not qMatricula.FieldByName("DATAEXPPASSAPORTE").IsNull Then
			CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime = qMatricula.FieldByName("DATAEXPPASSAPORTE").AsDateTime
		End If

		CurrentQuery.FieldByName("ORGAOEMISSOR").AsString = qMatricula.FieldByName("ORGAOEMISSORPASSAPORTE").AsString

	ElseIf Not qMatricula.FieldByName("DOCUMENTOIDENTIFICACAO").IsNull Then
		CurrentQuery.FieldByName("DOCUMENTO").AsString = qMatricula.FieldByName("DOCUMENTOIDENTIFICACAO").AsString
		CurrentQuery.FieldByName("TIPODOCUMENTO").AsString = "3"

		If Not qMatricula.FieldByName("DATAEXPEDICAODOCIDENTIFICACAO").IsNull Then
			CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime = qMatricula.FieldByName("DATAEXPEDICAODOCIDENTIFICACAO").AsDateTime
		End If

		CurrentQuery.FieldByName("ORGAOEMISSOR").AsString = qMatricula.FieldByName("ORGAOEMISSORDOCIDENTIFICACAO").AsString
	End If

	CurrentQuery.FieldByName("HANDLEBENEF").AsInteger = qBeneficiario.FieldByName("HANDLE").AsInteger
	CurrentQuery.FieldByName("CONTRATO").AsInteger = qBeneficiario.FieldByName("CONTRATO").AsInteger
	CurrentQuery.FieldByName("FAMILIA").AsInteger = qBeneficiario.FieldByName("FAMILIA").AsInteger


	Set qBeneficiario = Nothing
	Set qMatricula = Nothing

End Sub

Public Sub PopularCamposParaInclusao(pHandleFamilia As Long)
	Dim qFamiliaContrato As Object
	Set qFamiliaContrato = NewQuery
	qFamiliaContrato.Active = False
	qFamiliaContrato.Add("SELECT FAMILIA, CONTRATO FROM SAM_FAMILIA WHERE HANDLE = :HANDLE ")
	qFamiliaContrato.ParamByName("HANDLE").AsInteger = pHandleFamilia
	qFamiliaContrato.Active = True

	CurrentQuery.FieldByName("CONTRATO").AsInteger = qFamiliaContrato.FieldByName("CONTRATO").AsInteger
	CurrentQuery.FieldByName("FAMILIA").AsInteger = qFamiliaContrato.FieldByName("FAMILIA").AsInteger
	CurrentQuery.FieldByName("TIPOBENEFICIARIO").AsString = "D"

	Set qFamiliaContrato = Nothing

End Sub


Public Function RetornarTipoParentesco(pHandleContratoTipoDependente As Long) As Long
	Dim qContrato As Object
	Set qContrato = NewQuery
	qContrato.Active = False
	qContrato.Add("SELECT TIPODEPENDENTE FROM SAM_CONTRATO_TPDEP WHERE HANDLE = :HANDLE ")
	qContrato.ParamByName("HANDLE").AsInteger = pHandleContratoTipoDependente
	qContrato.Active = True

	RetornarTipoParentesco = qContrato.FieldByName("TIPODEPENDENTE").AsInteger

	Set qContrato = Nothing

End Function

Public Function RetornarTipoBeneficiario(pTipoDependente) As String
	Dim qTipoDependente As Object
	Set qTipoDependente = NewQuery
	qTipoDependente.Active = False
	qTipoDependente.Add("SELECT GRUPODEPENDENTE FROM SAM_TIPODEPENDENTE WHERE HANDLE = :HANDLE ")
	qTipoDependente.ParamByName("HANDLE").AsInteger = pTipoDependente
	qTipoDependente.Active = True

	RetornarTipoBeneficiario = qTipoDependente.FieldByName("GRUPODEPENDENTE").AsString

	Set qTipoDependente = Nothing

End Function
