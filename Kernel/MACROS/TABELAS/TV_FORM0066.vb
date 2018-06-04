'HASH: 23D38DA51D2B6A80A9E93DAA817F7377
'#Uses "*bsShowMessage"
Option Explicit

Dim SAM_PARAMETROSPRESTADOR As BPesquisa


Public Sub CPFCNPJ_OnExit()
    If CurrentQuery.FieldByName("CPFCNPJ").AsString <> "" Then
    	'Crislei.Sorrilha SMS 108068 - A comparação do if estava errada. ela tem de ser feita de acordo com a pageindex que o usuario esta
    	If FISICAJURIDICA.PageIndex = 1 Then
			If Not IsValidCGC(CurrentQuery.FieldByName("CPFCNPJ").AsString) Then
	        	bsShowMessage("CNPJ inválido", "E")
	        	CPFCNPJ.SetFocus
      		End If
    	Else
			If Not IsValidCPF(CurrentQuery.FieldByName("CPFCNPJ").AsString) Then
	        	bsShowMessage("CPF inválido: " + CurrentQuery.FieldByName("CPFCNPJ").AsString, "E")
        		CPFCNPJ.SetFocus
      		End If
    	End If
    End If
End Sub

Public Sub FISICAJURIDICA_OnChanging(AllowChange As Boolean)
  CurrentQuery.FieldByName("CPFCNPJ").Clear
  If FISICAJURIDICA.PageIndex = 1 Then
    CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
	FISICAJURIDICA.ReadOnly = False
	SEXO.ReadOnly = False
  Else
    CurrentQuery.FieldByName("CPFCNPJ").Mask = "99\.999\.999\/9999\-99;0;_"
    SEXO.ReadOnly = True
  End If
End Sub

Public Sub TABLE_AfterCommitted()
	Dim SamPrestadorBLL As CSBusinessComponent

    Set SamPrestadorBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.SamPrestadorBLL, Benner.Saude.Prestadores.Business")
   	SamPrestadorBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	SamPrestadorBLL.Execute("VerificarSeExportaBennerHospitalar")
End Sub

Public Sub TABLE_AfterInsert()
    Dim qTipoAutoriz As BPesquisa
	Set qTipoAutoriz = NewQuery
	Set SAM_PARAMETROSPRESTADOR = NewQuery

	SAM_PARAMETROSPRESTADOR.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")
	SAM_PARAMETROSPRESTADOR.Active = True

	CurrentQuery.FieldByName("CATEGORIA").AsInteger           = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHACATEGORIA").AsInteger
	CurrentQuery.FieldByName("TIPOPRESTADOR").AsInteger       = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHATIPOPRESTADOR").AsInteger
	CurrentQuery.FieldByName("TIPOPRESTADORODONTO").AsInteger = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHATIPOPRESTODONTO").AsInteger
	CurrentQuery.FieldByName("TIPOPRESTADORAMBAS").AsInteger  = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHATIPOPRESTAMBAS").AsInteger
	CurrentQuery.FieldByName("SOLICITANTE").AsString          = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHASOLICITANTE").AsString
	CurrentQuery.FieldByName("EXECUTOR").AsString             = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHAEXECUTOR").AsString
	CurrentQuery.FieldByName("LOCALEXECUCAO").AsBoolean       = SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHALOCALEXECUCAO").AsBoolean


	If (SessionVar("TIPOPEG_PLE") = "T") Then   'Tratamento odontológico

		CurrentQuery.FieldByName("FORMACAOPRESTADOR").AsInteger = 2

	Else
		If (SessionVar("TIPOGUIA_PLE") = "T") Then  'Tratamento odontológico

			CurrentQuery.FieldByName("FORMACAOPRESTADOR").AsInteger = 2

		End If
	End If

	If (SessionVar("HANDLETIPOAUTORIZ_PLE") <> "") Then

		qTipoAutoriz.Add(" SELECT TISSTIPOSOLICITACAO FROM SAM_TIPOAUTORIZ WHERE HANDLE = :PTIPOAUTORIZ ")
		qTipoAutoriz.ParamByName("PTIPOAUTORIZ").AsInteger = CLng(SessionVar("HANDLETIPOAUTORIZ_PLE"))
		qTipoAutoriz.Active = True

		If (qTipoAutoriz.FieldByName("TISSTIPOSOLICITACAO").AsString = "3") Then
			If (SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHATIPOPRESTODONTO").AsInteger > 0) Then
				CurrentQuery.FieldByName("FORMACAOPRESTADOR").AsInteger = 2
			Else
				If (SAM_PARAMETROSPRESTADOR.FieldByName("LIVREESCOLHATIPOPRESTAMBAS").AsInteger > 0) Then
					CurrentQuery.FieldByName("FORMACAOPRESTADOR").AsInteger = 3
				End If
			End If
		End If

		Set qTipoAutoriz = Nothing
	End If

	Set SAM_PARAMETROSPRESTADOR = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	Set SAM_PARAMETROSPRESTADOR = NewQuery

	SAM_PARAMETROSPRESTADOR.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")
	SAM_PARAMETROSPRESTADOR.Active = True

	PRESTADOR.ReadOnly = IIf((SAM_PARAMETROSPRESTADOR.FieldByName("DIGITARPRESTADOR").AsString = "N"), True, False)
	PRESTADOR.ReadOnly = IIf((SAM_PARAMETROSPRESTADOR.FieldByName("TABPADRAOCODIGO").AsInteger = 3), True, False)
	PRESTADOR.ReadOnly = IIf((SAM_PARAMETROSPRESTADOR.FieldByName("TABPADRAOCODIGO").AsInteger = 4), False, True)

	If (SAM_PARAMETROSPRESTADOR.FieldByName("EDITARTIPOPLE").AsBoolean = True) Then
		TIPOPRESTADOR.ReadOnly = False
		TIPOPRESTADORODONTO.ReadOnly = False
		TIPOPRESTADORAMBAS.ReadOnly = False
	Else
		TIPOPRESTADOR.ReadOnly = True
		TIPOPRESTADORODONTO.ReadOnly = True
		TIPOPRESTADORAMBAS.ReadOnly = True
	End If

	Set SAM_PARAMETROSPRESTADOR = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
    	CurrentQuery.FieldByName("CPFCNPJ").Mask = ""
  	Else
    	'Como na inclusão assume-se como Física a máscara inicial será de CPF
    	CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
 	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SamPrestLivreEscolha As Object
	Dim vsMensagem As String
	Dim viResult As Long
	Set SAM_PARAMETROSPRESTADOR = NewQuery

	SAM_PARAMETROSPRESTADOR.Clear
	SAM_PARAMETROSPRESTADOR.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")
	SAM_PARAMETROSPRESTADOR.Active = True

	If ((SAM_PARAMETROSPRESTADOR.FieldByName("EXIGIRCNPJCPF").AsString = "S") And _
		(Len(CurrentQuery.FieldByName("CPFCNPJ").AsString) <= 0)) Then
		bsShowMessage("CNPJ/CPF obrigatório", "E")

		CanContinue = False

		Exit Sub
	End If

	CurrentQuery.FieldByName("CPFCNPJ").AsString = Replace(CurrentQuery.FieldByName("CPFCNPJ").AsString, ".", "")
	CurrentQuery.FieldByName("CPFCNPJ").AsString = Replace(CurrentQuery.FieldByName("CPFCNPJ").AsString, "-", "")
	CurrentQuery.FieldByName("CPFCNPJ").AsString = Replace(CurrentQuery.FieldByName("CPFCNPJ").AsString, "/", "")

    If CurrentQuery.FieldByName("CPFCNPJ").AsString <> "" Then
	    If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
	      If Not IsValidCPF(CurrentQuery.FieldByName("CPFCNPJ").AsString) Then
	        CanContinue = False
	        bsShowMessage("CPF inválido", "E")
	        Exit Sub
	      End If
	    Else
	      If Not IsValidCGC(CurrentQuery.FieldByName("CPFCNPJ").AsString) Then
	        CanContinue = False
	        bsShowMessage("CNPJ inválido", "E")
	        Exit Sub
	      End If
	    End If
	End If

	If (Len(CurrentQuery.FieldByName("NOME").AsString) <= 0) Then
		bsShowMessage("O Campo nome não está preenchido", "E")

		CanContinue = False

		Exit Sub
	End If

	If CurrentQuery.FieldByName("FISICAJURIDICA").IsNull Then
		bsShowMessage("O Campo pessoa não está preenchido", "E")

		CanContinue = False

		Exit Sub
	End If

	If (CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1) Then
		If CurrentQuery.FieldByName("SEXO").IsNull Then
			bsShowMessage("O Campo sexo não está preenchido", "E")

			CanContinue = False

			Exit Sub
		End If
	End If


    If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then

      Dim VerificaBeneficiarioAtivoMesmoCPF As BPesquisa
	  Set VerificaBeneficiarioAtivoMesmoCPF = NewQuery

      VerificaBeneficiarioAtivoMesmoCPF.Active = False
      VerificaBeneficiarioAtivoMesmoCPF.Clear
	  VerificaBeneficiarioAtivoMesmoCPF.Add(" SELECT BLOQUEARINCLUSAOBENEF   ")
      VerificaBeneficiarioAtivoMesmoCPF.Add("   FROM SAM_CATEGORIA_PRESTADOR ")
      VerificaBeneficiarioAtivoMesmoCPF.Add("  WHERE HANDLE = :CATEGORIA     ")
      VerificaBeneficiarioAtivoMesmoCPF.ParamByName("CATEGORIA").AsInteger = CurrentQuery.FieldByName("CATEGORIA").AsInteger
  	  VerificaBeneficiarioAtivoMesmoCPF.Active = True

	  If (Not VerificaBeneficiarioAtivoMesmoCPF.EOF) And (VerificaBeneficiarioAtivoMesmoCPF.FieldByName("BLOQUEARINCLUSAOBENEF").AsString = "S" ) Then

	    If VerificaCPFDuplicado(CurrentQuery.FieldByName("CATEGORIA").AsInteger, CurrentQuery.FieldByName("CPFCNPJ").AsString, "titular") Then

	      bsShowMessage("O CPF informado pertence a beneficiário titular do sistema", "E")
	      CanContinue = False
	      Exit Sub

	    Else

	      If VerificaCPFDuplicado(CurrentQuery.FieldByName("CATEGORIA").AsInteger, CurrentQuery.FieldByName("CPFCNPJ").AsString, "dependente") Then

	        If bsShowMessage("O CPF informado pertence a beneficiário dependente do sistema. Deseja contiuar mesmo assim?", "Q") = vbNo Then
	          If VisibleMode Then
	            CanContinue = False
	          End If
	          Exit Sub
	        End If

	      Else

	        If VerificaCPFDuplicado(CurrentQuery.FieldByName("CATEGORIA").AsInteger, CurrentQuery.FieldByName("CPFCNPJ").AsString, "usuario") Then

	          If bsShowMessage("O CPF informado pertence a um profissional vinculado a um usuário do sistema. Deseja continuar mesmo assim?", "Q") = vbNo Then
	            If VisibleMode Then
	              CanContinue = False
	            End If
	            Exit Sub
	          End If

	        End If

	      End If

	    End If

	  End If
	  Set VerificaBeneficiarioAtivoMesmoCPF = Nothing
	End If

	Set SamPrestLivreEscolha = CreateBennerObject("SAMPRESTLIVREESCOLHA.CadastraPrestador")

	viResult = SamPrestLivreEscolha.Exec(CurrentSystem, _
										 CurrentQuery.TQuery, _
										 SAM_PARAMETROSPRESTADOR.TQuery, _
										 vsMensagem)

	Select Case viResult
		Case -1
			bsShowMessage("Processo abortado pelo usuário", "E")

			CanContinue = False
		Case 1
			bsShowMessage(vsMensagem, "E")

			CanContinue = False
	End Select

	Set SamPrestLivreEscolha = Nothing
End Sub

Public Function VerificaCPFDuplicado(handleCategoriaPrestador As Integer, cpf As String, tipoValidacao As String) As Boolean

	Dim interface As CSEntityCall

	If tipoValidacao = "titular" Then
	    Set interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfBeneficiarioTitularIgualCpfInformado")

	ElseIf tipoValidacao = "dependente" Then
		Set interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfBeneficiarioDependenteIgualCpfInformado")

	ElseIf tipoValidacao = "usuario" Then
		Set interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfUsuarioIgualCpfInformado")

	End If

	interface.AddParameter(pdtInteger, handleCategoriaPrestador)
	interface.AddParameter(pdtString, cpf)

	VerificaCPFDuplicado = CBool(interface.Execute())

	Set interface = Nothing

End Function

