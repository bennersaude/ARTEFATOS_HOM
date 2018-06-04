'HASH: 75FE7F06ACCC04B0C1259988A6B690A4
'MACRO =SAM_PRESTADOR

'#Uses "*VerificaEmail"
'#Uses "*bsShowMessage"
'#Uses "*CredenciamentoDePrestadores"
'#Uses "*RegistrarLogAlteracao"

Option Explicit

Dim NaoFaturarGuiasAnterior As String
Dim vgComplementoMascara As String
Dim vgInsertProponente As Boolean
Dim vgHandleProponente As Long
Dim vgFilial As Long
Dim vgHandleEnderecoProponente As Long
Dim vgPrestadorMestre As Long
Dim vInscricaoInss As String

Function Replicar(Caracter As String, Quantidade As Integer) As String
	Dim Contador As Integer

	Replicar = ""

	For Contador = 1 To Quantidade Step 1
		Replicar = Replicar + "0"
	Next Contador 
End Function


Public Sub PreparaCodigoPrestador
	Dim vMascara As String
	Dim vAux As String
	Dim i As Long
	Dim PARAM As Object
	Set PARAM = NewQuery

	PARAM.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")

	PARAM.Active = True

	PRESTADOR.ReadOnly = IIf(PARAM.FieldByName("DIGITARPRESTADOR").AsString = "N", True, False)

	If PARAM.FieldByName("TABPADRAOCODIGO").AsInteger = 3 Then
		PRESTADOR.ReadOnly = True
	ElseIf PARAM.FieldByName("TABPADRAOCODIGO").AsInteger = 4 Then
		PRESTADOR.ReadOnly = False
	End If

	Set PARAM = Nothing
End Sub

Public Sub ValidaCodigoPrestador(CanContinue As Boolean)
	Dim vErroMsg As String
	Dim vPRESTADOR As String
	Dim vConselho As String
	Dim vSequencia As Long
	Dim PARAM As Object
	Set PARAM = NewQuery

	PARAM.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")

	PARAM.Active = True

	'se estiver parametrizado que exige CPF/CNPJ o campo é obrigatório exceto para prestadores livre-escolha
	If PARAM.FieldByName("EXIGIRCNPJCPF").AsString = "S" Then
		If(CurrentQuery.FieldByName("CPFCNPJ").IsNull) And (CurrentQuery.FieldByName("CATEGORIA").AsInteger <>PARAM.FieldByName("LIVREESCOLHACATEGORIA").AsInteger) Then
			bsShowMessage("CPF/CNPJ obrigatório", "E")
			CanContinue = False
			Exit Sub
		End If
	End If

	'05/03/2002 -Fßbio
	'Independente do tipo do código deve validar as informaþ§es do conselho definidas nos parÔmetros gerais.
	If Len(CONSELHOREGIONAL.Text + UFCR.Text + CurrentQuery.FieldByName("REGIAOCR").AsString + CurrentQuery.FieldByName("INSCRICAOCR").AsString)>0 Then
		CanContinue = CheckConselho(PARAM.FieldByName("PADRAOCONSELHO").AsInteger, vErroMsg)

		If CanContinue = False Then
			bsShowMessage(vErroMsg, "E")
			Exit Sub
		End If
	End If

	Select Case PARAM.FieldByName("TABPADRAOCODIGO").AsInteger
		Case 0 'ERRO NO CËDIGO NOS PARAM#METROS DO  PRESTADOR
			CanContinue = False
			bsShowMessage("Erro na configuração do código.", "E")
			Set PARAM = Nothing
			Exit Sub
		Case 1 'CPF/CNPJ,conselho

			vPRESTADOR = Replace(Replace(Replace(CurrentQuery.FieldByName("CPFCNPJ").AsString, ".", ""), "/", ""), "-", "")

			If vPRESTADOR = "" And _
			   (Not CurrentQuery.FieldByName("CONSELHOREGIONAL").IsNull Or _
			    Not CurrentQuery.FieldByName("UFCR").IsNull Or _
			    Not CurrentQuery.FieldByName("REGIAOCR").IsNull Or _
			    Not CurrentQuery.FieldByName("INSCRICAOCR").IsNull) Then

				vPRESTADOR = FormataConselho(PARAM.FieldByName("PADRAOCONSELHO").AsInteger, PARAM.FieldByName("SEPARADORCONSELHO").AsString)

				'digitou um conselho mas Ú invßlido
				If vPRESTADOR <>"" Then
					CanContinue = CheckConselho(PARAM.FieldByName("PADRAOCONSELHO").AsInteger, vErroMsg)

					If Not CanContinue Then
						bsShowMessage(vErroMsg, "E")
						Exit Sub
					End If
				End If
			End If

			If vPRESTADOR = "" Then vPRESTADOR = CurrentQuery.FieldByName("PRESTADOR").AsString

			If PARAM.FieldByName("SOBREPORALTERACAO").AsString = "S" Then
				CurrentQuery.FieldByName("PRESTADOR").Value = vPRESTADOR
			Else
				If vPRESTADOR = "" Then
					bsShowMessage("Informar CPF/CNPJ ou o conselho.", "E")
					CanContinue = False
					Set PARAM = Nothing
					Exit Sub
				End If
			End If
		Case 2 'Conselho,CPF/CNPJ
			CanContinue = CheckConselho(PARAM.FieldByName("PADRAOCONSELHO").AsInteger, vErroMsg)

			If Not CanContinue Then
				bsShowMessage(vErroMsg, "E")
				Exit Sub
			End If

            If (Not CurrentQuery.FieldByName("CONSELHOREGIONAL").IsNull Or _
			    Not CurrentQuery.FieldByName("UFCR").IsNull Or _
			    Not CurrentQuery.FieldByName("REGIAOCR").IsNull Or _
			    Not CurrentQuery.FieldByName("INSCRICAOCR").IsNull) Then
  		      vPRESTADOR = FormataConselho(PARAM.FieldByName("PADRAOCONSELHO").AsInteger, PARAM.FieldByName("SEPARADORCONSELHO").AsString)
  		    End If

			If vPRESTADOR = "" Then vPRESTADOR = Replace(Replace(Replace(CurrentQuery.FieldByName("CPFCNPJ").AsString, ".", ""), "/", ""), "-", "")

			If vPRESTADOR = "" Then vPRESTADOR = CurrentQuery.FieldByName("PRESTADOR").AsString

			If PARAM.FieldByName("SOBREPORALTERACAO").AsString = "S" Then
				CurrentQuery.FieldByName("PRESTADOR").Value = vPRESTADOR
			Else
				If vPRESTADOR = "" Then
					bsShowMessage("Informar CPF/CNPJ ou o conselho.", "E")
					CanContinue = False
					Set PARAM = Nothing
					Exit Sub
				End If
			End If
		Case 3 'C¾digo automßtico
			If CurrentQuery.State = 3 Then
				NewCounter("SAM_PRESTADOR", 1, 1, vSequencia)
				CurrentQuery.FieldByName("PRESTADOR").Value = vSequencia
				CurrentQuery.FieldByName("PRESTADORNUMERICO").Value = vSequencia
			End If

			If Len(CurrentQuery.FieldByName("PRESTADOR").AsString)>Len(PARAM.FieldByName("MASCARAPRESTADOR").AsString) Then
				bsShowMessage("Prestador com mais dígitos do que permitido na máscara.", "E")
				CanContinue = False
				Set PARAM = Nothing
				Exit Sub
			End If
		Case 4 'C¾digo manual
			If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
				bsShowMessage("Informar PRESTADOR.", "E")
				CanContinue = False
				Set PARAM = Nothing
				Exit Sub
			End If

			If Len(CurrentQuery.FieldByName("PRESTADOR").AsString)>Len(PARAM.FieldByName("MASCARAPRESTADOR").AsString) Then
				bsShowMessage("Prestador com mais dígitos do que permitido na máscara.", "E")
				CanContinue = False
				Set PARAM = Nothing
				Exit Sub
			End If
	End Select

	Set PARAM = Nothing

	If CurrentQuery.FieldByName("PRESTADOR").IsNull Then
		bsShowMessage("Prestador obrigatório.", "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Function FormataConselho(pPadraoConselho As Long, pSeparador As String)As String
	Dim vConselho As String

	vConselho = BIN(pPadraoConselho)

	If Mid(vConselho, 3, 1) = 1 Then FormataConselho = CONSELHOREGIONAL.Text + pSeparador

	If Mid(vConselho, 2, 1) = 1 Then
		If Len(UFCR.Text)>0 Then FormataConselho = FormataConselho + UFCR.Text + pSeparador

		If Len(CurrentQuery.FieldByName("REGIAOCR").AsString)>0 Then FormataConselho = FormataConselho + CurrentQuery.FieldByName("REGIAOCR").AsString + pSeparador
	End If

	If Mid(vConselho, 1, 1) = 1 Then FormataConselho = FormataConselho + CurrentQuery.FieldByName("INSCRICAOCR").AsString + pSeparador

	If pSeparador <>"" Then FormataConselho = Mid(FormataConselho, 1, Len(FormataConselho) -1)
End Function

Public Function CheckConselho(pPadraoConselho As Long, vErroMsg As String)As Boolean
  CheckConselho = True
  'If pPadraoConselho = 0 Then
  '  CheckConselho = False
  '  vErroMsg = "'Padrão conselho' não foi informado nos parâmetros gerais de prestadores!"
  'Else
  If pPadraoConselho > 0 Then
	Dim vConselho As String

	vConselho = BIN(pPadraoConselho)
	CheckConselho = True
	vErroMsg = "Falta informar "

	If Mid(vConselho, 3, 1) = "1" Then
		If(CurrentQuery.FieldByName("CONSELHOREGIONAL").IsNull)Then
			CheckConselho = False
			vErroMsg = vErroMsg + " entidade,"
		End If
	End If

	If Mid(vConselho, 2, 1) = "1" Then
		If(CurrentQuery.FieldByName("UFCR").IsNull)And(CurrentQuery.FieldByName("REGIAOCR").IsNull)Then
			CheckConselho = False
			vErroMsg = vErroMsg + " estado ou região,"
		End If

		If(Not CurrentQuery.FieldByName("UFCR").IsNull)And(Not CurrentQuery.FieldByName("REGIAOCR").IsNull)Then
			CheckConselho = False
			vErroMsg = "Não informar estado e região do conselho simultaneamente."
			Exit Function
		End If
	End If

	If Mid(vConselho, 1, 1) = "1" Then
		If(CurrentQuery.FieldByName("INSCRICAOCR").IsNull)Then
			CheckConselho = False
			vErroMsg = vErroMsg + " inscrição,"
		End If
	End If

	vErroMsg = Mid(vErroMsg, 1, Len(vErroMsg) -1) + " do conselho."
  End If
End Function

Function BIN(P As Long)As String
	Dim i As Long
	Dim X As Long
	Dim zerosesquerda As String

	If P = 0 Then
		BIN = "0"
		Exit Function
	End If

	i = P
	zerosesquerda = ""

	While i >0
		X = i Mod 2
		i = Int(i / 2)

		If X = 0 Then BIN = "0" + BIN

		If X = 1 Then BIN = "1" + BIN

		zerosesquerda = zerosesquerda + "0"
	Wend

	BIN = Format(BIN, zerosesquerda)
End Function

Public Sub BOTAOCRIARUSUARIOWEB_OnClick()
' Alterada em 28/01/2010 - SMS 127892 - Evandro Zeferino
  Dim dll As Object
  Dim SQL As Object
  Dim vbValidaOperadora As Boolean
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT PV.VALIDACAOOPERADORA FROM TIS_PARAMETROSVALIDACAO PV JOIN TIS_PARAMETROS P ON PV.HANDLE = P.VALIDACAOPADRAO")
  SQL.Active = True
  vbValidaOperadora = SQL.FieldByName("VALIDACAOOPERADORA").AsString = "S"
  Set SQL = Nothing

  If Not vbValidaOperadora Then
    Dim vsMensagemErro As String
    Dim viResult As Integer

    Set dll = CreateBennerObject("BSPRE001.CriaUsuario")
    vsMensagemErro = ""

    viResult = dll.Criausuario(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "", 0, 0, vsMensagemErro)

    If viResult = 0 Then
      bsShowMessage("Usuário incluído/atualizado com sucesso.", "I")
    ElseIf viResult = 1 Then
      bsShowMessage(vsMensagemErro, "I")
    End If

  Else
    Set dll = CreateBennerObject("BSPRECRIAUSUARIO.CRIAUSUARIO")
    dll.CriaUsuario(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  End If

  Set dll = Nothing

End Sub

Public Sub BOTAOFINANCEIRO_OnClick()
	Dim Interface As Object
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Clear

	SQL.Add("SELECT HANDLE FROM SFN_CONTAFIN WHERE PRESTADOR=" + CurrentQuery.FieldByName("HANDLE").AsString)

	SQL.Active = True

	If Not SQL.EOF Then
		Set Interface = CreateBennerObject("SamContaFinanceira.Consulta")

		Interface.Exec(CurrentSystem, SQL.FieldByName("HANDLE").AsInteger)

		Set Interface = Nothing
	Else
		bsShowMessage("Conta financeira não encontrada", "E")
	End If

	Set SQL = Nothing
End Sub

Public Sub CBO_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCampos As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Set Interface = CreateBennerObject("Procura.Procurar")

	ShowPopup = False
	vTabela = "SAM_CBO"
	vColunas = "SAM_CBO.ESTRUTURA|SAM_CBO.DESCRICAO"
	vCriterio = "SAM_CBO.ULTIMONIVEL = 'S'"
	vCampos = "Estrutura|Descrição|Descrição --> Nível superior"
	vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "CBO", True, "")

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CBO").AsInteger = vHandle
	End If

	Set Interface = Nothing
End Sub
Public Sub CODIGOSERVICOPREFSP_OnPopup(ShowPopup As Boolean)

    Dim Interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String
    Dim vTabela  As String
    Dim vTitulo As String

	If CODIGOSERVICOPREFSP.PopupCase <> 0 Then
		ShowPopup = False
		Set Interface = CreateBennerObject("Procura.Procurar")

		vCampos = "Código|Descrição Abreviada"
		vColunas = "CODIGO|DESCRICAOABREVIADA"
		vTabela = "SFN_CODSERVICOS"
		vTitulo = "Códigos de Serviços Município São Paulo"
		vCriterio = ""

		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, vTitulo, True, CODIGOSERVICOPREFSP.LocateText)

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("CODIGOSERVICOPREFSP").AsInteger = vHandle
		CurrentQuery.FieldByName("CODIGOSERVICO").Value = Null
	End If
End Sub

'juliana criado em 14/06/2002 botÒo para mostrar detalhes de prestador.
Public Sub CODIGOSERVICO_OnPopup(ShowPopup As Boolean)
	Dim qVerificaConsiderarSp As BPesquisa
    Dim Interface As Object
    Dim vHandle As Long
    Dim vCampos As String
    Dim vColunas As String
    Dim vCriterio As String
    Dim vTabela  As String
    Dim vTitulo As String

   Set qVerificaConsiderarSp = NewQuery

   qVerificaConsiderarSp.Active = False
   qVerificaConsiderarSp.Clear
   qVerificaConsiderarSp.Add("SELECT CONSIDERARCODSERVICO FROM SFN_PARAMETROSFIN")
   qVerificaConsiderarSp.Active = True

    If (qVerificaConsiderarSp.FieldByName("CONSIDERARCODSERVICO").AsString = "S") Then
      If CurrentQuery.FieldByName("CODIGOSERVICOPREFSP").IsNull Then
        bsShowMessage("Necessário preencher o Cógido do Serviço do município de São Paulo.", "E")
        ShowPopup = False
        Exit Sub
      Else
        vCriterio = "(SAM_LISTASERVICOS.HANDLE IN (SELECT LISTASERVICO FROM SFN_CODSERVICOS_SERVICOSRELAC WHERE CODIGOSERVICO = " + CStr(CurrentQuery.FieldByName("CODIGOSERVICOPREFSP").AsInteger) + "))"
      End If
    Else
      vCriterio = ""
    End If

	If CODIGOSERVICO.PopupCase <> 0 Then
		ShowPopup = False
		Set Interface = CreateBennerObject("Procura.Procurar")

		vCampos = "Código|Código Exportação|Descrição"
		vColunas = "CODIGO|CODIGOEXPORTACAO|DESCRICAOABREVIADA"
		vTabela = "SAM_LISTASERVICOS"
		vTitulo = "Lista de serviços"


		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, vTitulo, True, CODIGOSERVICO.LocateText)

		Set Interface = Nothing
		Set qVerificaConsiderarSp = Nothing
	Else
		ShowPopup = True
	End If

	If vHandle <> 0 Then
	  CurrentQuery.Edit
	  CurrentQuery.FieldByName("CODIGOSERVICO").AsInteger = vHandle
    End If

End Sub
Public Sub DETALHESPRESTADOR_OnClick()
	Dim Interface As Object
	Set Interface = CreateBennerObject("CA005.ConsultaPrestador")

	Interface.info(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	Set Interface = Nothing
End Sub

Public Sub FISICAJURIDICA_OnChanging(AllowChange As Boolean)
  CurrentQuery.FieldByName("CPFCNPJ").Clear
  If FISICAJURIDICA.PageIndex = 1 Then
    CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
  Else
    CurrentQuery.FieldByName("CPFCNPJ").Mask = "99\.999\.999\/9999\-99;0;_"
  End If
End Sub

Public Sub ITEMNFRPA_OnChange()
	If (Not (CurrentQuery.FieldByName("ITEMNFRPA").IsNull) And CurrentQuery.FieldByName("ITEMNFRPA").AsInteger = CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").AsInteger) Then
		bsShowMessage("O Item NF e o Item NF Internação, não podem ter o mesmo valor.", "E")
		CurrentQuery.FieldByName("ITEMNFRPA").Value = Null
		Exit Sub
	End If
End Sub

Public Sub ITEMNFRPAINTERNACAO_OnChange()
	If (Not (CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").IsNull) And CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").AsInteger = CurrentQuery.FieldByName("ITEMNFRPA").AsInteger) Then
		bsShowMessage("O Item NF Internação e o Item NF, não podem ter o mesmo valor.", "E")
		CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").Value = Null
		Exit Sub
	End If
End Sub

Public Sub TABLE_AfterCommitted()
    If VisibleMode Then
	  Dim SamPrestadorProcBLL As CSBusinessComponent
	  Dim SamPrestadorBLL As CSBusinessComponent
 	  Dim retorno As Boolean

	  Set SamPrestadorBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.SamPrestadorBLL, Benner.Saude.Prestadores.Business")
   	  SamPrestadorBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	  SamPrestadorBLL.Execute("VerificarSeExportaBennerHospitalar")

      Set SamPrestadorProcBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
      SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAVANCADO")
      retorno = SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")

      If (retorno) Then
		 SamPrestadorProcBLL.ClearParameters()
         SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAUTOMATICO")
         retorno = SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")

          If (retorno) Then
			 SamPrestadorProcBLL.ClearParameters()
             SamPrestadorProcBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
             retorno = SamPrestadorProcBLL.Execute("VerificarExistenciaDeProcessoDeCredenciamentoNoPrestador")

             If (Not retorno) Then
               If MsgBox("Deseja incluir processo de credenciamento para o prestador?",vbYesNo)=vbYes Then
                BOTAOINICIARCREDENCIAMENTO_OnClick
               End If
             End If
		 End If
      End If
    End If
	Set SamPrestadorProcBLL = Nothing
	Set SamPrestadorBLL     = Nothing
End Sub

Public Sub TABLE_AfterDelete()
    If CurrentQuery.FieldByName("RECEBEDOR").AsString = "S" Then
      Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
      Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

      TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
      TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_PRESTADOR")
      TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "Z")

      TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
    End If
End Sub

Public Sub TABLE_AfterEdit()
  If (CurrentQuery.FieldByName("TABEMISSAOAUTOMATICARPA").IsNull) Then
	CurrentQuery.FieldByName("TABEMISSAOAUTOMATICARPA").AsInteger = 1
  End If
End Sub

Public Sub TABLE_AfterInsert()
  If VisibleMode Then
    CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
  End If
  If WebMode Then
  	If Not (SessionVar("HANDLE_PROPONENTE") = "") Then
	  vgHandleProponente = CLng(SessionVar("HANDLE_PROPONENTE"))
  	Else
  		If(CurrentEntity.TransitoryVars("PROPONENTE").IsPresent)Then
  			vgHandleProponente = CurrentEntity.TransitoryVars("PROPONENTE").AsInteger
  		End If
  	End If

	ImportarProponente
	SessionVar("HANDLE_PROPONENTE") = ""
  End If
End Sub

Public Sub TABLE_AfterPost()
	RegistrarLogAlteracao "SAM_PRESTADOR", CurrentQuery.FieldByName("HANDLE").AsInteger, "TABLE_AfterPost"

	Dim SQL As Object
	Dim QryInss As Object
	Dim vFilial As Long
	Dim vLog As String
	Dim Cadastro As Object
	Dim qConsulta As Object
	Dim vPrestadorMestre As Long
	Dim qPrestador As Object
	Dim qFechamento As Object
	Dim qVerifica As Object
	Dim qCadastro As Object
	Dim vbControleQuestao As Boolean
	Set SQL = NewQuery
	Set QryInss = NewQuery

    Dim Interface As Object
    Set Interface = CreateBennerObject("FINANCEIRO.ContaFin")
    If Interface.Cadastro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 2, 0)<0 Then
      bsShowMessage("Erro ao criar conta financeira", "E")
      Exit Sub
    End If
    Set Interface = Nothing

	'insere dados do proponente
	If vgInsertProponente = True Then
		INSERTENDERECO
		INSERTCURRICULO
		INSERTCURRICULOEXPERIENCIA
		UPDATEPROPONENTE
		INSERTESPECIALIZADES

		'O Start foi dado na insercao do prestador feito no botao importacao do proponente
		vgInsertProponente = False

        If VisibleMode Then
		  RefreshNodesWithTable("SAM_PROPONENTE")
		  RefreshNodesWithTable("SAM_PRESTADOR")
		End If
	End If

	' se for selecionado municipiopagamento o prestador fica com filial padrao =filial da regiao do municipio
	' a checagem já é feita na macro do especifico
	If CurrentQuery.FieldByName("DATADESCREDENCIAMENTO").IsNull Then
		If (SessionVar("ALTERAESPEC") = "") And (Not CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").IsNull) Then

			If CurrentQuery.FieldByName("FILIALPADRAO").AsInteger <>vgFilial Then
				If vgFilial > 0 Then
					SQL.Active = False

					SQL.Clear

					SQL.Add("UPDATE SAM_PRESTADOR SET FILIALPADRAO = :FILIAL WHERE HANDLE = :PRESTADOR")

					SQL.ParamByName("FILIAL").Value = vgFilial
					SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

					SQL.ExecSQL
				End If
			End If
		End If

		If((Not CurrentQuery.FieldByName("ESTADOPAGAMENTO").IsNull)And(CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").IsNull))Then

			If CurrentQuery.FieldByName("FILIALPADRAO").AsInteger <>vgFilial Then
				If vgFilial > 0 Then
					SQL.Active = False

					SQL.Clear

					SQL.Add("UPDATE SAM_PRESTADOR SET FILIALPADRAO = :FILIAL WHERE HANDLE = :PRESTADOR")

					SQL.ParamByName("FILIAL").Value = vgFilial
					SQL.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

					SQL.ExecSQL
				End If
			End If
		End If
	End If

	'******************************* Alterado em 07/11/2002 Durval implantação da SMS 13557 ***********************************************
	'**************************************************************************************************************************************
	Set qVerifica = NewQuery

	qVerifica.ForceNoLockOnTables = False

	qVerifica.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")

	qVerifica.Active = True

	If qVerifica.FieldByName("UTILIZACADASTROINVERTIDO").AsString = "S" Then
		If vgPrestadorMestre <>0 And CurrentQuery.FieldByName("PRESTADORMESTRE").IsNull Then
			Set qPrestador = NewQuery

			qPrestador.Add("SELECT NOME ")
			qPrestador.Add("  FROM SAM_PRESTADOR")
			qPrestador.Add(" WHERE HANDLE = :HANDLE")

			qPrestador.ParamByName("HANDLE").Value = vgPrestadorMestre
			qPrestador.Active = True

			'Em modo web não é possível ter questões no meio do código, pois neste ambiente
			'a macro é executado duas vezes para tratar questões, sendo que na primeira execução
			'a resposta é sempre não.
			vbControleQuestao = True

			If VisibleMode Then
				If bsShowMessage("A vigência do prestador será fechada como membro do grupo empresarial " + (Chr(13)) + _
				   qPrestador.FieldByName("NOME").AsString + " ?", "Q") = vbNo Then
					vbControleQuestao = False
				End If
			End If

			If vbControleQuestao Then
				Set qFechamento = NewQuery

				qFechamento.Add("UPDATE SAM_PRESTADOR_PRESTADORDAENTID")
				qFechamento.Add("   SET DATAFINAL = :DATAHOJE")
				qFechamento.Add(" WHERE PRESTADOR = :PRESTADOR")
				qFechamento.Add("   AND ENTIDADE  = :ENTIDADE")

				qFechamento.ParamByName("DATAHOJE").Value = ServerDate
				qFechamento.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
				qFechamento.ParamByName("ENTIDADE").Value = vgPrestadorMestre

				qFechamento.ExecSQL

				If WebMode Then
					bsShowMessage("A vigência do prestador foi fechada como membro do grupo empresarial " + (Chr(13)) + _
				   				  qPrestador.FieldByName("NOME").AsString, "I")
				End If
			End If

			qPrestador.Active = False
		End If

		If Not CurrentQuery.FieldByName("PRESTADORMESTRE").IsNull Then
			Set qConsulta = NewQuery

			qConsulta.Add("SELECT HANDLE                                           ")
			qConsulta.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID                   ")
			qConsulta.Add(" WHERE PRESTADOR   = :PRESTADOR                         ")
			qConsulta.Add("   AND ENTIDADE    = :ENTIDADE                          ")
			qConsulta.Add("   AND DATAINICIAL <= :DATAHOJE                         ")
			qConsulta.Add("   AND ((DATAFINAL IS NULL) OR (DATAFINAL > :DATAHOJE)) ")

			qConsulta.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
			qConsulta.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger
			qConsulta.ParamByName("DATAHOJE").Value = ServerDate
			qConsulta.Active = True

			If qConsulta.EOF Then
				'Em modo web não é possível ter questões no meio do código, pois neste ambiente
				'a macro é executado duas vezes para tratar questões, sendo que na primeira execução
				'a resposta é sempre não.
				vbControleQuestao = True

				If bsShowMessage("Deseja cadastrar o prestador como membro de corpo clínico do grupo empresarial selecionado?", _
								 "Q") = vbNo Then
					vbControleQuestao = False
				End If

				If vbControleQuestao Then
                  'SMS 90453 - Barbosa - 09/05/2008
				  If VisibleMode Then
				    Set Cadastro = CreateBennerObject("BSPRE003.Rotinas")

					Cadastro.Cadastrar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger)

					Set Cadastro = Nothing
                  Else
                     Set qCadastro = NewQuery

	 				 qCadastro.Add("INSERT INTO SAM_PRESTADOR_PRESTADORDAENTID")
					 qCadastro.Add("            (HANDLE,         ")
					 qCadastro.Add("             DATAINICIAL,    ")
					 qCadastro.Add("             ENTIDADE,       ")
					 qCadastro.Add("             PRESTADOR,      ")
					 qCadastro.Add("             TEMPORARIO,     ")
					 qCadastro.Add("             TABPAGAMENTO,   ")
					 qCadastro.Add("             PRECO,          ")
					 qCadastro.Add("             ISENTOIRRF)     ")
					 qCadastro.Add("      VALUES                 ")
					 qCadastro.Add("            (:HANDLE,        ")
					 qCadastro.Add("             :DATAINICIAL,   ")
					 qCadastro.Add("             :ENTIDADE,      ")
					 qCadastro.Add("             :PRESTADOR,     ")
					 qCadastro.Add("             :TEMPORARIO,    ")
					 qCadastro.Add("             :TABPAGAMENTO,  ")
					 qCadastro.Add("             :PRECO,         ")
					 qCadastro.Add("             :ISENTOIRRF)    ")

					 qCadastro.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_PRESTADORDAENTID")
					 qCadastro.ParamByName("DATAINICIAL").Value = ServerDate
					 qCadastro.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger
					 qCadastro.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
					 qCadastro.ParamByName("TEMPORARIO").Value = "N"
					 qCadastro.ParamByName("TABPAGAMENTO").Value = "1"
					 qCadastro.ParamByName("PRECO").Value = "L"
					 qCadastro.ParamByName("ISENTOIRRF").Value = "N"

					 qCadastro.ExecSQL

                     bsShowMessage("Prestador cadastrado como membro de corpo clínico do grupo empresarial selecionado!", "I")
				  End If
				End If
			End If

			qConsulta.Active = False
		End If

		qVerifica.Active = False
	Else
		If vgPrestadorMestre <>0 And CurrentQuery.FieldByName("PRESTADORMESTRE").IsNull Then
			Set qFechamento = NewQuery

			qFechamento.Add("UPDATE SAM_PRESTADOR_PRESTADORDAENTID")
			qFechamento.Add("   SET DATAFINAL = :DATAHOJE")
			qFechamento.Add(" WHERE PRESTADOR = :PRESTADOR")
			qFechamento.Add("   AND ENTIDADE  = :ENTIDADE")

			qFechamento.ParamByName("DATAHOJE").Value = ServerDate
			qFechamento.ParamByName("PRESTADOR").Value = vgPrestadorMestre
			qFechamento.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

			qFechamento.ExecSQL
		End If

		Set qConsulta = NewQuery

		qConsulta.Add("SELECT HANDLE                                           ")
		qConsulta.Add("  FROM SAM_PRESTADOR_PRESTADORDAENTID                   ")
		qConsulta.Add(" WHERE PRESTADOR   = :PRESTADOR                         ")
		qConsulta.Add("   AND ENTIDADE    = :ENTIDADE                          ")
		qConsulta.Add("   AND DATAINICIAL <= :DATAHOJE                         ")
		qConsulta.Add("   AND ((DATAFINAL IS NULL) OR (DATAFINAL > :DATAHOJE)) ")

		qConsulta.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
		qConsulta.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger
		qConsulta.ParamByName("DATAHOJE").Value = ServerDate
		qConsulta.Active = True

		If qConsulta.EOF And(Not CurrentQuery.FieldByName("PRESTADORMESTRE").IsNull)Then
			Set qCadastro = NewQuery

			qCadastro.Add("INSERT INTO SAM_PRESTADOR_PRESTADORDAENTID")
			qCadastro.Add("            (HANDLE,         ")
			qCadastro.Add("             DATAINICIAL,    ")
			qCadastro.Add("             ENTIDADE,       ")
			qCadastro.Add("             PRESTADOR,      ")
			qCadastro.Add("             TEMPORARIO,     ")
			qCadastro.Add("             TABPAGAMENTO,   ")
			qCadastro.Add("             PRECO,          ")
			qCadastro.Add("             ISENTOIRRF)     ")
			qCadastro.Add("      VALUES                 ")
			qCadastro.Add("            (:HANDLE,        ")
			qCadastro.Add("             :DATAINICIAL,   ")
			qCadastro.Add("             :ENTIDADE,      ")
			qCadastro.Add("             :PRESTADOR,     ")
			qCadastro.Add("             :TEMPORARIO,    ")
			qCadastro.Add("             :TABPAGAMENTO,  ")
			qCadastro.Add("             :PRECO,         ")
			qCadastro.Add("             :ISENTOIRRF)    ")

			qCadastro.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_PRESTADORDAENTID")
			qCadastro.ParamByName("DATAINICIAL").Value = ServerDate
			qCadastro.ParamByName("ENTIDADE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
			qCadastro.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger
			qCadastro.ParamByName("TEMPORARIO").Value = "N"
			qCadastro.ParamByName("TABPAGAMENTO").Value = "1"
			qCadastro.ParamByName("PRECO").Value = "L"
			qCadastro.ParamByName("ISENTOIRRF").Value = "N"

			qCadastro.ExecSQL
		End If

		qConsulta.Active = False
		qVerifica.Active = False
	End If
	'**************************************************************************************************************************************
	'**************************************************************************************************************************************

	If Not CurrentQuery.FieldByName("INSCRICAOINSS").IsNull Then
		If vInscricaoInss <>CurrentQuery.FieldByName("INSCRICAOINSS").AsString Then
			QryInss.Active = False

			QryInss.Clear

			QryInss.Add("UPDATE SAM_PRESTADOR_INSS        ")
			QryInss.Add("   SET INSCRICAO = :PINSCRICAO ")
			QryInss.Add(" WHERE PRESTADOR = :PPRESTADOR ")

			QryInss.ParamByName("PINSCRICAO").Value = CurrentQuery.FieldByName("INSCRICAOINSS").AsString
			QryInss.ParamByName("PPRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

			QryInss.ExecSQL

			Set QryInss = Nothing
		End If
	End If

    If CurrentQuery.FieldByName("RECEBEDOR").AsString = "S" Then
      Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
      Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

      TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
      TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_PRESTADOR")
      TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

      TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
    End If
End Sub

Public Sub TABLE_AfterScroll()
	Dim qVerificaConsiderarSp As BPesquisa

    If VisibleMode Then
        BOTAOGERARESPECIALIDADES.Visible = False
        BOTAOCRIARUSUARIOWEB2.Visible = False
        BOTAOINICIARCREDENCIAMENTO.Visible = False
    End If

	Dim S As Object
	Set S = NewQuery

	S.Clear

	S.Add("SELECT TABTIPOGESTAO")
	S.Add("  FROM EMPRESAS E,")
	S.Add("       FILIAIS F")
	S.Add(" WHERE F.HANDLE = :FILIAL")
	S.Add("   AND E.HANDLE = F.EMPRESA")

	S.ParamByName("FILIAL").Value = CurrentQuery.FieldByName("FILIALPADRAO").AsInteger
	S.Active = True

	If VisibleMode = True Then
		If S.FieldByName("TABTIPOGESTAO").AsInteger = 3 Then
			COOPERATIVAORIGEM.Visible = True
		Else
			COOPERATIVAORIGEM.Visible = False
		End If
	End If

	PreparaCodigoPrestador

	BOTAOFINANCEIRO.Enabled = IIf(CurrentQuery.State = 1, True, False)

    If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
      CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
    Else
      CurrentQuery.FieldByName("CPFCNPJ").Mask = "99\.999\.999\/9999\-99;0;_"
    End If

	'se o CPF/CNPJ estiver preenchido nÒo pode alterar,independente dos parÔmetros do c¾digo.
	CPFCNPJ.ReadOnly = IIf(CurrentQuery.FieldByName("CPFCNPJ").IsNull, False, True)

	MostraAfastamento

  'Proponente e alterar categoria serÒo revistos.Esperando Garcia.
  '	BOTAOALTERARCATEGORIA.Enabled =False
  S.Active = False

  Set S = Nothing

  vInscricaoInss = CurrentQuery.FieldByName("INSCRICAOINSS").AsString
	'===================   SMS 70314 25/10/2006 DURVAL ===============================================================================
	'Rodrigo
	If WebMode Then
		SessionVar("EVENTOS") =   "A.HANDLE IN (SELECT EX.EVENTO " _
														+ "               FROM SAM_PRESTADOR_ESPECIALIDADE  PE                                              				" _
														+ "               JOIN SAM_ESPECIALIDADE      E ON (E.HANDLE = PE.ESPECIALIDADE)                    				" _
														+ "               JOIN SAM_ESPECIALIDADEGRUPO EG ON (E.HANDLE = EG.ESPECIALIDADE)                   				" _
														+ "               JOIN SAM_ESPECIALIDADEGRUPO_EXEC  EX ON (EX.ESPECIALIDADEGRUPO = EG.HANDLE)       				" _
														+ "              WHERE PE.PRESTADOR = "+CurrentQuery.FieldByName("HANDLE").AsString+"               				" _
														+ "                and not EXISTS (SELECT X.HANDLE                                                  				" _
														+ "                                  FROM SAM_PRESTADOR_ESPECIALIDADEGRP X                          				" _
														+ "                                 WHERE (X.PRESTADORESPECIALIDADE = PE.HANDLE)      								" _
														+ "                                   and (PE.DATAINICIAL <=" + SQLDate(ServerDate)+ " ) 								" _
														+ "                                   and (PE.DATAFINAL is NULL or PE.DATAFINAL >=" + SQLDate(ServerDate)+ " )  		" _
														+ "                                   and (PE.PRESTADOR = "+CurrentQuery.FieldByName("HANDLE").AsString+")) 		" _
														+ "             UNION      																							" _
														+ "             SELECT EX.EVENTO      																				" _
														+ "               FROM SAM_PRESTADOR_ESPECIALIDADE    PE      														" _
														+ "               JOIN SAM_ESPECIALIDADE              E  ON (E.HANDLE = PE.ESPECIALIDADE)      						" _
														+ "               JOIN SAM_PRESTADOR_ESPECIALIDADEGRP PG ON (PG.PRESTADORESPECIALIDADE = PE.HANDLE)      			" _
														+ "               JOIN SAM_ESPECIALIDADEGRUPO         EG ON (PE.ESPECIALIDADE = EG.ESPECIALIDADE            		" _
														+ "                                                           and PG.ESPECIALIDADEGRUPO = EG.HANDLE)      			" _
														+ "               JOIN SAM_ESPECIALIDADEGRUPO_EXEC    EX ON (EX.ESPECIALIDADEGRUPO = EG.HANDLE)     				" _
														+ "              WHERE (PE.DATAINICIAL <= " + SQLDate(ServerDate)+ " )													" _
														+ "                and (PE.DATAFINAL is NULL or PE.DATAFINAL >=" + SQLDate(ServerDate)+ " ) 							" _
														+ "                and (PG.DATAINICIAL <= " + SQLDate(ServerDate)+ " ) 													" _
														+ "                and (PG.DATAFINAL is NULL or PG.DATAFINAL >= " + SQLDate(ServerDate)+ " ) 							" _
														+ "                and (PE.PRESTADOR = " +CurrentQuery.FieldByName("HANDLE").AsString+") 							" _
														+ "                and EXISTS (SELECT X.HANDLE          															" _
														+ "                              FROM SAM_PRESTADOR_ESPECIALIDADEGRP X          									" _
														+ "                             WHERE X.PRESTADORESPECIALIDADE = PE.HANDLE)      									" _
														+ "             UNION      																							" _
														+ "             SELECT PR.EVENTO      																				" _
														+ "               FROM SAM_PRESTADOR_REGRA PR      																	" _
														+ "              WHERE (PR.PRESTADOR = "+CurrentQuery.FieldByName("HANDLE").AsString+") 							" _
														+ "                and (PR.REGRAEXCECAO = 'R')      																" _
														+ "                and (PR.DATAINICIAL <= " + SQLDate(ServerDate)+ " ) 													" _
														+ "                and (PR.DATAFINAL is NULL or PR.DATAFINAL >= " + SQLDate(ServerDate)+ " ) 						)	" _
														+ "                and A.HANDLE not IN (SELECT PRX.EVENTO       														" _
														+ "                                       FROM SAM_PRESTADOR_REGRA PRX       										" _
														+ "                                      WHERE (PRX.PRESTADOR = "+CurrentQuery.FieldByName("HANDLE").AsString+") 	" _
														+ "                                        and (PRX.REGRAEXCECAO = 'E')       										" _
														+ "                                        and (PRX.DATAINICIAL <= " + SQLDate(ServerDate)+ " )					 	   	" _
														+ "                                        and (PRX.DATAFINAL is NULL or PRX.DATAFINAL >= " + SQLDate(ServerDate)+ " )  	" _
														+ "and A.ULTIMONIVEL = 'S')                                                                                    "
	End If
	'=======================================================================================================================================

  If WebMode Then
    If WebMenuCode = "T3879" Then
	  TIPOPRESTADOR.ReadOnly = True
    End If
  End If

  Set qVerificaConsiderarSp = NewQuery

  qVerificaConsiderarSp.Active = False
  qVerificaConsiderarSp.Clear
  qVerificaConsiderarSp.Add("SELECT CONSIDERARCODSERVICO FROM SFN_PARAMETROSFIN")
  qVerificaConsiderarSp.Active = True

  If (qVerificaConsiderarSp.FieldByName("CONSIDERARCODSERVICO").AsString = "N")  Then
    CODIGOSERVICOPREFSP.Visible = False
  Else
    CODIGOSERVICOPREFSP.Visible = True
  End If

  Set qVerificaConsiderarSp = Nothing

  If VisibleMode Then
  	If (CurrentQuery.FieldByName("DATAPRIMEIROCREDEN").IsNull And CurrentQuery.State <> 3) Then

	  Dim SamPrestadorProcBLL As CSBusinessComponent
 	  Dim retorno As Boolean

      Set SamPrestadorProcBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
      SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAVANCADO")
      retorno = SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")

        If (retorno) Then
		   SamPrestadorProcBLL.ClearParameters()
           SamPrestadorProcBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
           retorno = SamPrestadorProcBLL.Execute("VerificarExistenciaDeProcessoDeCredenciamentoNoPrestador")

           If (Not retorno) Then
              BOTAOINICIARCREDENCIAMENTO.Visible = True
           End If
		End If
	Set SamPrestadorProcBLL = Nothing
    End If
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String
	Dim SQL As Object
	Dim TemPermissao As Boolean

	TemPermissao = True

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		TemPermissao = False
	End If

	If TemPermissao Then
		Set SQL = NewQuery

		'Se nao existe fatura, excluir a contafin e o prestador
		SQL.Clear

		SQL.Add("SELECT HANDLE FROM SAM_PROPONENTE WHERE PRESTADOR = :HPRESTADOR")

		SQL.ParamByName("HPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		SQL.Active = True

		If SQL.EOF Then
			SQL.Clear

			SQL.Add("DELETE FROM SFN_CONTAFIN")
			SQL.Add(" WHERE PRESTADOR = :HPRESTADOR")
			SQL.Add("   AND NOT EXISTS (SELECT HANDLE")
			SQL.Add("                   FROM SFN_FATURA")
			SQL.Add("                  WHERE CONTAFINANCEIRA = SFN_CONTAFIN.HANDLE)")

			SQL.ParamByName("HPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			SQL.ExecSQL
		Else
			Dim MsgProp As String

			MsgProp = "O prestador foi incluído através da importação de proponente, portanto não pode ser excluído!"

			bsShowMessage(MsgProp, "E")
		End If
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim vFiltro As String
	Dim vFilial As String
	Dim vFiltroFilialE As String
	Dim vFiltroFilialM As String
	Dim Msg As String
	Dim vUsuario As String

	vFiltro = checkPermissaoFilial(CurrentSystem, "A", "P", Msg)
	'******* Alterado 07/11/2002 ***** SMS 13557 *****************************
	vgPrestadorMestre = CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger
	'*************************************************************************

	If vFiltro = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	vUsuario = Str(CurrentUser)

	If VisibleMode = True Then
		'se estiver abaixo da carga de filiais filtra os estados/municipios daquela filial +controle de acesso
		vFiltroFilialE = ""
		vFiltroFilialM = ""

		If RecordHandleOfTable("FILIAIS")>0 Then '  Or(vFilial <>0)Then
			vFilial = Str(RecordHandleOfTable("FILIAIS"))
			vFiltroFilialE = " AND ESTADOS.HANDLE IN (SELECT ESTADO FROM FILIAIS_ESTADOS WHERE FILIAL = " + vFilial + " )"
			vFiltroFilialM = " AND MUNICIPIOS.REGIAO IN (SELECT HANDLE FROM SAM_REGIAO WHERE FILIAL = " + vFilial + " )"
		End If
	End If

	'um municipio no qual o usuario possui permissao pode estar em um estado no qual nÒo possui
	UpdateLastUpdate("ESTADOS")
	UpdateLastUpdate("MUNICIPIOS")

	If VisibleMode Then
		MUNICIPIOPAGAMENTO.LocalWhere = " MUNICIPIOS.HANDLE IN " + vFiltro + vFiltroFilialM
		ESTADOPAGAMENTO.LocalWhere =" ESTADOS.HANDLE IN (SELECT M.ESTADO " + _
																									 "   FROM Z_GRUPOUSUARIOS_FILIAIS GUF, " + _
																									 "        MUNICIPIOS M, " + _
																									 "        SAM_REGIAO R " + _
																									 "  WHERE GUF.USUARIO = " +vUsuario + _
																									 "    AND M.REGIAO = R.HANDLE " + _
																									 "    AND GUF.FILIAL = R.FILIAL " + _
																									 "    AND GUF.ALTERAR = 'S' " + _
																									 "  UNION" + _
																									 " SELECT M.ESTADO" + _
																									 "   FROM Z_GRUPOUSUARIOS GUF," + _
																									 "        MUNICIPIOS M, " + _
																									 "        SAM_REGIAO R " + _
																									 " WHERE GUF.HANDLE = " +vUsuario + _
																									 "   AND M.REGIAO = R.HANDLE " + _
																									 "   AND GUF.FILIALPADRAO = R.FILIAL ) "
	Else
		MUNICIPIOPAGAMENTO.WebLocalWhere = " A.HANDLE IN " + vFiltro + vFiltroFilialM
		ESTADOPAGAMENTO.WebLocalWhere =" A.HANDLE IN (SELECT M.ESTADO " + _
																								"   FROM Z_GRUPOUSUARIOS_FILIAIS GUF, " + _
																								"        MUNICIPIOS M, " + _
																								"        SAM_REGIAO R " + _
																								"  WHERE GUF.USUARIO = " +vUsuario + _
																								"    AND M.REGIAO = R.HANDLE " + _
																								"    AND GUF.FILIAL = R.FILIAL " + _
																								"    AND GUF.ALTERAR = 'S' " + _
																								"  UNION" + _
																								" SELECT M.ESTADO" + _
																								"   FROM Z_GRUPOUSUARIOS GUF," + _
																								"        MUNICIPIOS M, " + _
																								"        SAM_REGIAO R " + _
																								" WHERE GUF.HANDLE = " +vUsuario + _
																								"   AND M.REGIAO = R.HANDLE " + _
																								"   AND GUF.FILIALPADRAO = R.FILIAL ) "
	End If

	If VisibleMode = True Then
		ISS.LocalWhere = "FISICAJURIDICA = @FISICAJURIDICA"
	Else
		ISS.WebLocalWhere = "FISICAJURIDICA = @CAMPO(FISICAJURIDICA)"
		CBO.WebLocalWhere = "A.ULTIMONIVEL = 'S'"
	End If

	NaoFaturarGuiasAnterior = CurrentQuery.FieldByName("NAOFATURARGUIAS").AsString
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
    If WebMode Then
      CurrentQuery.FieldByName("CPFCNPJ").Mask = ""
    Else
      'Como na inclusão assume-se como Física a máscara inicial será de CPF
      If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
        CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
      Else
        CurrentQuery.FieldByName("CPFCNPJ").Mask = "99\.999\.999\/9999\-99;0;_"
      End If
    End If

	Dim vFiltro As String
	Dim vFilial As String
	Dim vFiltroFilialE As String
	Dim vFiltroFilialM As String
	Dim Msg As String
	Dim vUsuario As String
	Dim vMunicipio As String
	Dim qVerifica As Object

	vFiltro = checkPermissaoFilial(CurrentSystem, "I", "P", Msg)

	If vFiltro = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If VisibleMode Then
		MUNICIPIOPAGAMENTO.LocalWhere = " MUNICIPIOS.HANDLE IN " + vFiltro + vFiltroFilialM
		ESTADOPAGAMENTO.LocalWhere =" ESTADOS.HANDLE IN (SELECT M.ESTADO " + _
																									 "   FROM Z_GRUPOUSUARIOS_FILIAIS GUF, " + _
																									 "        MUNICIPIOS M, " + _
																									 "        SAM_REGIAO R " + _
																									 "  WHERE GUF.USUARIO = " +Str(CurrentUser) + _
																									 "    AND M.REGIAO = R.HANDLE " + _
																									 "    AND GUF.FILIAL = R.FILIAL " + _
																									 "    AND GUF.ALTERAR = 'S' " + _
																									 "  UNION" + _
																									 " SELECT M.ESTADO" + _
																									 "   FROM Z_GRUPOUSUARIOS GUF," + _
																									 "        MUNICIPIOS M, " + _
																									 "        SAM_REGIAO R " + _
																									 " WHERE GUF.HANDLE = " +Str(CurrentUser) + _
																									 "   AND M.REGIAO = R.HANDLE " + _
																									 "   AND GUF.FILIALPADRAO = R.FILIAL ) "
	Else
		MUNICIPIOPAGAMENTO.WebLocalWhere = " A.HANDLE IN " + vFiltro + vFiltroFilialM
		ESTADOPAGAMENTO.WebLocalWhere =" A.HANDLE IN (SELECT M.ESTADO " + _
																											"   FROM Z_GRUPOUSUARIOS_FILIAIS GUF, " + _
																											"        MUNICIPIOS M, " + _
																											"        SAM_REGIAO R " + _
																											"  WHERE GUF.USUARIO = " +Str(CurrentUser) + _
																											"    AND M.REGIAO = R.HANDLE " + _
																											"    AND GUF.FILIAL = R.FILIAL " + _
																											"    AND GUF.ALTERAR = 'S' " + _
																											"  UNION" + _
																											" SELECT M.ESTADO" + _
																											"   FROM Z_GRUPOUSUARIOS GUF," + _
																											"        MUNICIPIOS M, " + _
																											"        SAM_REGIAO R " + _
																											" WHERE GUF.HANDLE = " +Str(CurrentUser) + _
																											"   AND M.REGIAO = R.HANDLE " + _
																											"   AND GUF.FILIALPADRAO = R.FILIAL ) "
	End If

	vUsuario = Str(CurrentUser)
	vMunicipio = Str(CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").AsInteger)

	If (VisibleMode = True) And (Not WebMode) Then
		'se estiver abaixo da carga de filiais filtra os estados/municipios daquela filial +controle de acesso
		If RecordHandleOfTable("FILIAIS")>0 Then
			vFilial = Str(RecordHandleOfTable("FILIAIS"))
			vFiltroFilialE = "AND HANDLE IN (SELECT ESTADO FROM FILIAIS_ESTADOS WHERE FILIAL = " + vFilial + " )"
			vFiltroFilialM = "AND REGIAO IN (SELECT HANDLE FROM SAM_REGIAO WHERE FILIAL = " + vFilial + " )"
		Else
			vFiltroFilialE = ""
			vFiltroFilialM = ""
		End If
	End If

	'um municipio no qual o usuario possui permissao pode estar em um estado no qual nÒo possui
	UpdateLastUpdate("ESTADOS")
	UpdateLastUpdate("MUNICIPIOS")

	If VisibleMode = True Then
		ISS.LocalWhere = "FISICAJURIDICA = @FISICAJURIDICA"
	Else
		ISS.WebLocalWhere = "FISICAJURIDICA = @CAMPO(FISICAJURIDICA)"
	End If

	NaoFaturarGuiasAnterior = "N"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


	PreparaCodigoPrestador

	If (CurrentQuery.FieldByName("TABEMISSAOAUTOMATICARPA").AsInteger = 2) Then

		If (CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 2) Or (CurrentQuery.FieldByName("EMITENFPOR").AsInteger = 3) Then
			bsShowMessage("A emissão do RPA disponível somente para pessoa física que emite nota fiscal.", "E")
			CanContinue = False
			Exit Sub
		End If
	End If

	Dim qVerifica As Object
	Dim qFilial As Object
	Dim ParametrosProcContas As BPesquisa
	Dim NaoVerificarPoisNaoEhLivreEscolha As String
	Dim PodeDuplicarCpfCnpj As String 'SMS 103725 - Rafael Zarpellon - 13/10/2008
	Set qVerifica = NewQuery
	Set qFilial = NewQuery
	Set ParametrosProcContas = NewQuery

	If VisibleMode = True Then
		qVerifica.ForceNoLockOnTables = False
	End If

	If (SessionVar("ALTERAESPEC") = "") And (Not CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").IsNull) Then
		qFilial.Active = False
		qFilial.Clear
		qFilial.Add("SELECT R.FILIAL FROM SAM_REGIAO R, MUNICIPIOS M WHERE M.REGIAO = R.HANDLE AND M.HANDLE = :MUNICIPIO")
		qFilial.ParamByName("MUNICIPIO").Value = CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").AsString
		qFilial.Active = True

		If qFilial.EOF Then
			bsShowMessage("Região do município de pagamento não possui filial parametrizada!", "E")
			CanContinue = False
			Exit Sub
		Else
			vgFilial = qFilial.FieldByName("FILIAL").AsInteger
		End If
	End If

	If((Not CurrentQuery.FieldByName("ESTADOPAGAMENTO").IsNull)And(CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").IsNull))Then
		qFilial.Active = False
		qFilial.Clear
		qFilial.Add("SELECT FILIAL FROM FILIAIS_ESTADOS WHERE ESTADO = :ESTADO")
		qFilial.ParamByName("ESTADO").Value = CurrentQuery.FieldByName("ESTADOPAGAMENTO").AsString
		qFilial.Active = True

		If qFilial.EOF Then
			bsShowMessage("Região do estado de pagamento não possui filial parametrizada!", "E")
			CanContinue = False
			Exit Sub
		Else
			vgFilial = qFilial.FieldByName("FILIAL").AsInteger
		End If
	End If

	qVerifica.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")

	qVerifica.Active = True

	NaoVerificarPoisNaoEhLivreEscolha = "N"

	PodeDuplicarCpfCnpj = qVerifica.FieldByName("CHKCPFCNPJPLE").AsString 'SMS 103725 - Rafael Zarpellon - 13/10/2008

	If qVerifica.FieldByName("LIVREESCOLHACATEGORIA").AsString <>"" Then
		If CurrentQuery.FieldByName("CATEGORIA").AsString <>qVerifica.FieldByName("LIVREESCOLHACATEGORIA").AsString Then
			NaoVerificarPoisNaoEhLivreEscolha = "S"
		End If
	End If

	If NaoVerificarPoisNaoEhLivreEscolha = "N" Then
		If qVerifica.FieldByName("CHKCPFCNPJPLE").AsString = "S" Then 'Verifica CNPJ/CPF do prestador
			qVerifica.Clear

			qVerifica.Add("SELECT COUNT(*) QTD FROM SAM_PRESTADOR WHERE CPFCNPJ = :CPFCNPJ AND HANDLE <> :HANDLE")

			qVerifica.ParamByName("CPFCNPJ").Value = CurrentQuery.FieldByName("CPFCNPJ").AsString
			qVerifica.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
			qVerifica.Active = True

			If qVerifica.FieldByName("QTD").AsInteger >0 Then
				bsShowMessage("Já existe um prestador com este CPF/CNPJ cadastrado", "E")
				CanContinue = False
			End If
		End If

		Set qVerifica = Nothing
	End If

	Dim SQL As Object
	Set SQL = NewQuery
	Dim QueryPeso As Object
	Set QueryPeso = NewQuery

	'juliana alterado em 31/05/2002 para verificar alteraþÒo no motivo de referenciamento e consequentemente alteraþÒo do peso
	If CurrentQuery.State = 2 Then
		QueryPeso.Clear

		QueryPeso.Add("SELECT M.PESO FROM SAM_MOTIVOREFERENCIAMENTO M WHERE M.HANDLE=:MOTIVOREF")

		QueryPeso.ParamByName("MOTIVOREF").Value = CurrentQuery.FieldByName("MOTIVOREFERENCIAMENTO").AsInteger
		QueryPeso.Active = False
		QueryPeso.Active = True

		'se o motivo de referenciamento for nulo.
		If CurrentQuery.FieldByName("MOTIVOREFERENCIAMENTO").IsNull Then
			CurrentQuery.FieldByName("PESO").Clear
		End If

		If Not(CurrentQuery.FieldByName("MOTIVOREFERENCIAMENTO").IsNull)Then
			CurrentQuery.FieldByName("PESO").AsInteger = QueryPeso.FieldByName("PESO").AsInteger
		End If

		Set QueryPeso = Nothing
	End If

	If ((Not CurrentQuery.FieldByName("REGISTROOPERINTERMEDIARIO").IsNull) And (Len(CurrentQuery.FieldByName("REGISTROOPERINTERMEDIARIO").AsString) <> 6)) Then
		bsShowMessage("Informe um regitro da operadora intermediária com 6 caracteres.", "E")
		CanContinue = False
		Exit Sub
	End If

	If Len(CurrentQuery.FieldByName("NOME").AsString)<3 Then
		bsShowMessage("Informe um nome com mais que três caracteres.", "E")
		CanContinue = False
		Set SQL = Nothing
		Exit Sub
	End If


      If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then

          Dim vRetorno As String

	      Dim ValidarESocial As CSBusinessComponent
	      Set ValidarESocial = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.eSocial.eSocialBLL, Benner.Saude.Prestadores.Business")

	      ValidarESocial.AddParameter(pdtString, CurrentQuery.FieldByName("PIS").AsString)
	      vRetorno = ValidarESocial.Execute("ValidarPIS")

	      If vRetorno <> "" Then
			bsShowMessage(vRetorno, "E")
			CanContinue = False
			Set ValidarESocial = Nothing
			Exit Sub
		  End If


          If CurrentQuery.FieldByName("RECEBEDOR").AsString = "S" Then

			  ValidarESocial.ClearParameters
			  ValidarESocial.AddParameter(pdtString, CurrentQuery.FieldByName("Z_NOME").AsString)

			  vRetorno = ValidarESocial.Execute("ValidarNome")

			   If vRetorno <> "" Then
				bsShowMessage(vRetorno, "I")
			   End If
          End If

          Set ValidarESocial = Nothing

      End If



	'se for recebedor o CPF/CNPJ Ú obrigat¾rio,independente dos parÔmetros gerais exigir CPF/CNPJ
	If (CurrentQuery.FieldByName("RECEBEDOR").AsString = "S" Or _
	    CurrentQuery.FieldByName("EXECUTOR").AsString = "S") Then
		If CurrentQuery.FieldByName("CPFCNPJ").IsNull Then
			bsShowMessage("CPF/CNPJ obrigatório para recebedores ou executores.", "E")
			CanContinue = False
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	If (CurrentQuery.FieldByName("RECEBEDOR").AsString  = "S") Then
		If(CurrentQuery.FieldByName("ESTADOPAGAMENTO").IsNull)Or(CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").IsNull)Then
			bsShowMessage("Estado e município de pagamento obrigatório para recebedores.", "E")
			CanContinue = False
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	'valida CPF/CNPJ se informado
    If Not CurrentQuery.FieldByName("CPFCNPJ").IsNull Then
      If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
        If Not IsValidCPF(CurrentQuery.FieldByName("CPFCNPJ").AsString) Then
          CanContinue = False
          bsShowMessage("CPF inválido", "E")
          Set SQL = Nothing
          Exit Sub
        End If
      Else
        If Not IsValidCGC(CurrentQuery.FieldByName("CPFCNPJ").AsString) Then
          CanContinue = False
          bsShowMessage("CNPJ inválido", "E")
          Set SQL = Nothing
          Exit Sub
        End If
      End If
	End If

	If CurrentQuery.FieldByName("NOMELIVRO").Value = "F" Then
		If CurrentQuery.FieldByName("FANTASIA").IsNull Then
			bsShowMessage("Nome Fantasia obrigatório.", "E")
			CanContinue = False
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("EMAIL").IsNull Then
		If Not VerificaEmail(CurrentQuery.FieldByName("EMAIL").AsString)Then
			bsShowMessage("E-mail inválido", "E")
			CanContinue = False
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	If (Not CurrentQuery.FieldByName("EMAILAUTORIZACAO").IsNull) And (CurrentQuery.FieldByName("PADRAORESPOSTA").AsString = "E") Then
		If Not VerificaEmail(CurrentQuery.FieldByName("EMAILAUTORIZACAO").AsString)Then
			bsShowMessage("E-mail de autorizações inválido", "E")
			CanContinue = False
			Set SQL = Nothing
			Exit Sub
		End If
	End If

	' Checar rede diferenciada
	If CurrentQuery.FieldByName("REDEDIFERENCIADA").AsString = "N" Then
		Dim q1 As Object
		Set q1 = NewQuery

		q1.Add("SELECT handle FROM SAM_REDEDIFERENCIADA_PRESTADOR WHERE PRESTADOR = :handle")

		q1.ParamByName("handle").Value = CurrentQuery.FieldByName("handle").AsInteger
		q1.Active = True

		If Not q1.EOF Then
			CanContinue = False
			bsShowMessage("O prestador possui rede diferenciada cadastrada.", "E")
			q1.Active = False
			Set q1 = Nothing
			Set SQL = Nothing
			Exit Sub
		End If

		Set q1 = Nothing
	End If

    'SMS 87115 - Marcelo Barbosa - 11/09/2007
	If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
		If CurrentQuery.FieldByName("INSCRICAOINSS").AsString <>"" Then
			If Not IsNumeric(CurrentQuery.FieldByName("INSCRICAOINSS").AsString) Then
				bsShowMessage("Número de Inscrição não pode conter letras. Somente Números e ponto.", "I")
				CanContinue = False
			End If
		End If
	End If
	'Fim - SMS 87115

	'consistir tipo de ISS com Fis/Jur
	SQL.Add("SELECT * FROM SFN_ISS WHERE HANDLE = :ISS")

	SQL.ParamByName("ISS").Value = CurrentQuery.FieldByName("ISS").AsInteger
	SQL.Active = True

	If Not SQL.EOF Then
		If(SQL.FieldByName("FISICAJURIDICA").AsInteger = 1 And CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger <>1)Then
			Set SQL = Nothing
			CanContinue = False
			bsShowMessage("Tipo de ISS permitido somente para Pessoa Física!", "E")
			Exit Sub
		End If

		If(SQL.FieldByName("FISICAJURIDICA").AsInteger = 2 And CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger <>2)Then
			Set SQL = Nothing
			CanContinue = False
			bsShowMessage("Tipo de ISS permitido somente para Pessoa Jurídica!", "E")
			Exit Sub
		End If
	End If

	If (Not CurrentQuery.FieldByName("RG").IsNull) And _
		 Trim(CurrentQuery.FieldByName("RG").AsString)<>CurrentQuery.FieldByName("RG").AsString Then
		bsShowMessage("RG Inválido", "E")
		CanContinue = False
		Set SQL = Nothing
		Exit Sub
	End If

	'verifcacao de datas
	If CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime >ServerDate Then
		CanContinue = False
		bsShowMessage("Data de inclusão não pode ser maior que a data atual.", "E")
		Set SQL = Nothing
		Exit Sub
	End If

	If CurrentQuery.FieldByName("DATANASCIMENTO").AsDateTime >ServerDate Then
		CanContinue = False
		bsShowMessage("Data de nascimento não pode ser maior que a data atual.", "E")
		Set SQL = Nothing
		Exit Sub
	End If

	If CurrentQuery.FieldByName("DATAINSCRICAOCR").AsDateTime >ServerDate Then
		CanContinue = False
		bsShowMessage("Data de inclusão no Conselho Regional não pode ser maior que a data atual.", "E")
		Set SQL = Nothing
		Exit Sub
	End If

    'se retirar o motivo do bloqueio limpa a data
    If CurrentQuery.FieldByName("MOTIVOBLOQUEIO").IsNull Then
      CurrentQuery.FieldByName("DATABLOQUEIO").Value = Null
    ElseIf CurrentQuery.FieldByName("DATABLOQUEIO").IsNull Then
      CurrentQuery.FieldByName("DATABLOQUEIO").Value = ServerDate
    End If

	'limpar dados do tab tipoprestador que nÒo estiver selecionado
	If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
		CurrentQuery.FieldByName("INSCRICAOESTADUAL").Clear
		CurrentQuery.FieldByName("CORPOCLINICO").Value = "N"
		CurrentQuery.FieldByName("ASSOCIACAO").Value = "N"
		CurrentQuery.FieldByName("COOPERATIVA").Value = "N"
	Else
		CurrentQuery.FieldByName("DATANASCIMENTO").Clear
		CurrentQuery.FieldByName("ESTADOCIVIL").Clear
		CurrentQuery.FieldByName("SEXO").Clear
		CurrentQuery.FieldByName("RG").Clear
		CurrentQuery.FieldByName("ORGAOEMISSOR").Clear
		CurrentQuery.FieldByName("ESTADO").Clear
		CurrentQuery.FieldByName("NATURALIDADE").Clear
		CurrentQuery.FieldByName("NACIONALIDADE").Clear
		CurrentQuery.FieldByName("CENTRALPAGER").Clear
		CurrentQuery.FieldByName("PAGER").Clear
		CurrentQuery.FieldByName("CELULAR").Clear
	End If

	If CurrentQuery.State = 2 Then
		If NaoFaturarGuiasAnterior <>CurrentQuery.FieldByName("NAOFATURARGUIAS").AsString Then
			Dim texto As String

			If CurrentQuery.FieldByName("NAOFATURARGUIAS").AsString = "S" Then
				texto = "O Prestador foi modificado para NÃO faturar guias ! "
			Else
				texto = "O Prestador foi modificado para faturar guias ! "
			End If

			bsShowMessage(texto, "I")
		End If
	End If

	'valida o c¾digo do prestador conforme o tipo de c¾digo escolhido nos parÔmetros gerais.
	ValidaCodigoPrestador(CanContinue)

	If CanContinue = False Then
		Exit Sub
		Set SQL = Nothing
	End If
	If PodeDuplicarCpfCnpj = "S" Then '--este IF foi incluído por claudemir em  09/12/02 -sms 13323 'SMS 103725 - Rafael Zarpellon - 13/10/2008 - Incluído And PodeDuplicarCpfCnpj = "S"
		SQL.Clear
		SQL.Add("SELECT HANDLE FROM SAM_PRESTADOR WHERE CPFCNPJ = :CPFCNPJ ")
		If CurrentQuery.State = 2 Then SQL.Add("AND HANDLE <> :HANDLE")
		SQL.ParamByName("CPFCNPJ").Value = CurrentQuery.FieldByName("CPFCNPJ").AsString
		If CurrentQuery.State = 2 Then SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
		SQL.Active = True

		If Not SQL.EOF Then
			CanContinue = False
			bsShowMessage("CPF/CNPJ já cadastrado.", "E")
			Exit Sub
			Set SQL = Nothing
		End If

		Set SQL = Nothing
	End If

	If VisibleMode Then
		If RecordHandleOfTable("FILIAIS")>0 Then
			CurrentQuery.FieldByName("FILIALPADRAO").Value = RecordHandleOfTable("FILIAIS")
		End If
	End If

	If Not CurrentQuery.FieldByName("CONSELHOREGIONAL").IsNull Then
		Dim S As Object
		Set S = NewQuery

		S.Add("SELECT COUNT(HANDLE) C")
		S.Add("  FROM SAM_PRESTADOR P")
		S.Add(" WHERE P.HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString)
		S.Add("   AND P.CONSELHOREGIONAL = " + CurrentQuery.FieldByName("CONSELHOREGIONAL").AsString)
		S.Add("   AND P.INSCRICAOCR = '" + CurrentQuery.FieldByName("INSCRICAOCR").AsString + "'")

		If CurrentQuery.FieldByName("UFCR").IsNull Then
			S.Add("AND P.UFCR IS NULL")
		Else
			S.Add("AND P.UFCR = " + CurrentQuery.FieldByName("UFCR").AsString)
		End If

		If CurrentQuery.FieldByName("REGIAOCR").IsNull Then
			S.Add("AND P.REGIAOCR IS NULL")
		Else
			S.Add("AND P.REGIAOCR = '" + CurrentQuery.FieldByName("REGIAOCR").AsString + "'")
		End If

		S.Active = True

		If S.FieldByName("C").AsInteger >0 Then
			bsShowMessage("Existe outro registro com a mesma inscrição de conselho regional!", "I")
		End If

		S.Active = False

		Set S = Nothing
	End If

	'SMS 26917 -Roger Cazangi -20/08/2004 -Início
	Dim qParamFin As Object
	Set qParamFin = NewQuery

	qParamFin.Add("SELECT UTILIZACENTROCUSTO, BUSCACENTROCUSTOPAGAMENTO")
	qParamFin.Add("FROM SFN_PARAMETROSFIN")

	qParamFin.Active = True

	If(qParamFin.FieldByName("UTILIZACENTROCUSTO").AsString = "S")And(qParamFin.FieldByName("BUSCACENTROCUSTOPAGAMENTO").AsString <>"C")Then
		If(CurrentQuery.FieldByName("FILIALCUSTO").IsNull)Then
			bsShowMessage("A filial de custo deve ser informada devido à busca de centro de custo pela filial do prestador.", "E")
			CanContinue = False
		End If
	End If

	Set qParamFin = Nothing
	'SMS 26917 -Roger Cazangi -20/08/2004 -Fim

	'SMS 43212 - 25.01.2006 - Consistindo campos do Convênio de Reciprocidade
	If ((CurrentQuery.FieldByName("CONVENIORECIPROCIDADE").AsString = "S") And (CurrentQuery.FieldByName("RECEBEDOR").AsString = "N")) Then
		If (VisibleMode) Then
			bsShowMessage("Prestador não é Recebedor. Não pode ser marcado como Convênio de Reciprocidade.", "E")
			CanContinue = False
		Else
			CurrentQuery.FieldByName("CONVENIORECIPROCIDADE").AsString = "N"
		End If
	End If

	If (CurrentQuery.FieldByName("CONVENIORECIPROCIDADE").AsString = "N") And _
	   (CurrentQuery.FieldByName("ISENTAAUTORIZ").AsString = "S") Then
		bsShowMessage("Prestador não é Convênio de Reciprocidade, o campo 'Isenta Autorização' foi desmarcado!", "I")

		CurrentQuery.FieldByName("ISENTAAUTORIZ").AsString = "N"
	End If
	'Final SMS 43212

	Dim PARAM As Object
	Set PARAM = NewQuery

	PARAM.Add("SELECT * FROM SAM_PARAMETROSPRESTADOR")

	PARAM.Active = True

	If CurrentQuery.State > 1 Then
		' Julio - SMS : 71778 - 13/11/2006
		' Se o padrão de código for AUTOMÁTICO ou MANUAL
		If ((PARAM.FieldByName("TABPADRAOCODIGO").AsInteger = 3) Or _
				(PARAM.FieldByName("TABPADRAOCODIGO").AsInteger = 4)) Then
			If (PARAM.FieldByName("ZEROMASCARA").AsString = "S") Then
				Dim QtdMk As Integer
				Dim QtdPt As Integer
				Dim Qtd As Integer

				QtdMk = Len(PARAM.FieldByName("MASCARAPRESTADOR").AsString)
				QtdPt = Len(CurrentQuery.FieldByName("PRESTADOR").AsString)

				If QtdPt < QtdMk Then
					Qtd = QtdMk - QtdPt

					CurrentQuery.FieldByName("PRESTADOR").AsString = Replicar("0",Qtd) + CurrentQuery.FieldByName("PRESTADOR").AsString
				End If
			End If
		End If
		' Julio - SMS : 71778 - Fim
	End If

	Dim qVerificaPrestador As BPesquisa
	Set qVerificaPrestador = NewQuery

	qVerificaPrestador.Add(" SELECT PRESTADOR              ")
    qVerificaPrestador.Add("   FROM SAM_PRESTADOR          ")
    qVerificaPrestador.Add("  WHERE PRESTADOR = :PRESTADOR ")
    qVerificaPrestador.Add("    AND HANDLE <> :HANDLE      ")
    qVerificaPrestador.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsString
    qVerificaPrestador.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	qVerificaPrestador.Active = True

	If Not qVerificaPrestador.EOF Then
       bsShowMessage("Prestador já Cadastrado!", "E")
       CanContinue = False
       Set qVerificaPrestador = Nothing
	   Exit Sub
	End If

    Set qVerificaPrestador = Nothing

	If (Not CurrentQuery.FieldByName("CODIGOSERVICOPREFSP").IsNull) And (CurrentQuery.FieldByName("CODIGOSERVICO").IsNull) And (CODIGOSERVICOPREFSP.Visible = True) Then
	  bsShowMessage("Necessário informar um Código de Serviço relacionado com a Pref. de São Paulo.", "E")
      CanContinue = False
      Exit Sub
	End If

	If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then

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

	        End If

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

    If (CurrentQuery.FieldByName("EMITENFPOR").OldValue <> CurrentQuery.FieldByName("EMITENFPOR").NewValue) Then
		Dim qPagamento As BPesquisa
      	Set qPagamento = NewQuery

	    qPagamento.Add("SELECT COUNT(HANDLE) QUANTIDADE ")
    	qPagamento.Add("  FROM SAM_AGRUPADORPAGAMENTO ")
    	qPagamento.Add(" WHERE RECEBEDOR = :HANDLEPRESTADOR ")
    	qPagamento.Add("   AND STATUSPAGAMENTO <> '4' ")
    	qPagamento.ParamByName("HANDLEPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    	qPagamento.Active = True

    	If (qPagamento.FieldByName("QUANTIDADE").AsInteger > 0) Then
			CanContinue = False
			bsShowMessage("Não é permitido alterar o campo 'Emissão de nota fiscal por' quando o prestador possui pagamento pendente de faturamento.", "E")
			Set qPagamento = Nothing
			Exit Sub
		End If
		Set qPagamento = Nothing
	End If

	If ((CurrentQuery.FieldByName("ISENTANF").AsString = "S") And (CurrentQuery.FieldByName("EMITENFPOR").AsInteger <> 3)) Then
		CanContinue = False
		bsShowMessage("Campo 'Emissão de nota fiscal por' deve ser configurado com a opção 'Não apresenta' quando o parâmetro 'Isenta apresentação de nota fiscal' estiver marcado.", "E")
		Exit Sub
	End If

	If Not (((CurrentQuery.FieldByName("VALORMINIMOMEDIA").AsFloat = 0) And (CurrentQuery.FieldByName("QUANTIDADEPAGAMENTOS").AsFloat = 0) And (CurrentQuery.FieldByName("PODEEXCEDERMEDIA").AsFloat = 0)) Or _
      ((CurrentQuery.FieldByName("VALORMINIMOMEDIA").AsFloat <> 0) And (CurrentQuery.FieldByName("QUANTIDADEPAGAMENTOS").AsFloat <> 0) And (CurrentQuery.FieldByName("PODEEXCEDERMEDIA").AsFloat <> 0)))	Then
		CanContinue = False
		bsShowMessage("Parametrização de 'Média de pagamento' incompleta.", "E")
		Exit Sub
	End If

	Dim mensagemRetorno As String
	If Not(ValidarRegraEmissaoAutomaticaRPA(mensagemRetorno)) Then
		CanContinue = False
		bsShowMessage(mensagemRetorno, "E")
		Exit Sub
	End If

	ParametrosProcContas.Active = False
	ParametrosProcContas.Add("SELECT * ")
	ParametrosProcContas.Add("  FROM SAM_PARAMETROSPROCCONTAS	")
	ParametrosProcContas.Active = True

	If(CurrentQuery.FieldByName("PREVIAIMPOSTOS").AsString = "S" And _
		(ParametrosProcContas.FieldByName("TABIMPOSTOSNANF").AsInteger = 2 Or ParametrosProcContas.FieldByName("TABCONTROLEPAGAMENTO").AsInteger = 2 Or ParametrosProcContas.FieldByName("TABCONCILIACAODOCFISCAIS").AsInteger = 2 Or _
		CurrentQuery.FieldByName("EMITENFPOR").AsInteger = 3)) Then
		bsShowMessage("Para o parâmetro 'Realizar prévia de Impostos' ser marcado o campo 'Emite NF por' deve estar marcado como 'Valor Líquido' ou 'Valor Bruto', e nos parâmetros gerais Processamento de Contas o campo 'Impostos na NF' deve estar marcado como 'Controlar'.","E")
		CurrentQuery.FieldByName("PREVIAIMPOSTOS").AsString = "N"
		Set ParametrosProcContas = Nothing
		Exit Sub
	End If
	Set ParametrosProcContas = Nothing
End Sub

Public Function ValidarRegraEmissaoAutomaticaRPA(ByRef mensagemRetorno As String) As Boolean

    mensagemRetorno = ""

	If (TABEMISSAOAUTOMATICARPA.PageIndex = 1) Then
		Dim qDotacaoOrcamentariaLigada As BPesquisa
		Set qDotacaoOrcamentariaLigada = NewQuery

		qDotacaoOrcamentariaLigada.Add("SELECT CONTROLADOTORC FROM SFN_PARAMETROSFIN ")
		qDotacaoOrcamentariaLigada.Active = True

		If (qDotacaoOrcamentariaLigada.FieldByName("CONTROLADOTORC").AsInteger = 2) Then
			If (Not CurrentQuery.FieldByName("ITEMNFRPA").IsNull) Then
				If Not(ValidarItensNFParaRPA(CurrentQuery.FieldByName("ITEMNFRPA").AsInteger)) Then

					ValidarRegraEmissaoAutomaticaRPA = False
					mensagemRetorno = "Para emissão automática de RPA escolher um item de imposto vigente e que esteja contido nas regras de ambos os impostos do prestador, de origem recurso orçamento e recurso próprio."

					qDotacaoOrcamentariaLigada.Active = False
					Set qDotacaoOrcamentariaLigada = Nothing
					Exit Function
				End If
			End If

			If (Not CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").IsNull) Then
				If Not(ValidarItensNFParaRPA(CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").AsInteger)) Then

					ValidarRegraEmissaoAutomaticaRPA = False
					mensagemRetorno = "Para emissão automatica de RPA escolher um item de internação de imposto vigente e que esteja contido nas regras de ambos os impostos do prestador, de origem recurso orçamento e recurso próprio."

					qDotacaoOrcamentariaLigada.Active = False
					Set qDotacaoOrcamentariaLigada = Nothing
					Exit Function
				End If
			End If

		Else
			If (Not CurrentQuery.FieldByName("ITEMNFRPA").IsNull) Then
				Dim qImposto As BPesquisa
      			Set qImposto = NewQuery

			    qImposto.Add("SELECT COUNT(1) EXITEIMPOSTOVIGENTE")
			    qImposto.Add("  FROM SAM_PRESTADOR_IMPOSTO PI ")
			    qImposto.Add("  JOIN SFN_IMPOSTO_ITEM ITEM ON ITEM.IMPOSTO = PI.IMPOSTO ")
		    	qImposto.Add(" WHERE (:DATA >= PI.DATAINICIAL AND (PI.DATAFINAL IS NULL OR :DATA <= PI.DATAFINAL)) ")
		    	qImposto.Add("   AND PI.PRESTADOR = :HANDLEPRESTADOR ")
		    	qImposto.Add("   AND ITEM.ITEM = :HANDLEITEMIMPOSTO ")
				qImposto.ParamByName("HANDLEPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
   		 		qImposto.ParamByName("HANDLEITEMIMPOSTO").AsInteger = CurrentQuery.FieldByName("ITEMNFRPA").AsInteger
   		 		qImposto.ParamByName("DATA").Value = ServerDate
    			qImposto.Active = True

		    	If (qImposto.FieldByName("EXITEIMPOSTOVIGENTE").AsInteger = 0) Then
		    	    ValidarRegraEmissaoAutomaticaRPA = False
					mensagemRetorno = "Para emissão automatica de RPA escolher um item de imposto vigente do prestador."
					Set qImposto = Nothing
					Exit Function
				End If

				Set qImposto = Nothing
			End If

			If (Not CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").IsNull) Then
    		  	Set qImposto = NewQuery

			    qImposto.Add("SELECT COUNT(1) EXITEIMPOSTOVIGENTE")
			    qImposto.Add("  FROM SAM_PRESTADOR_IMPOSTO PI ")
			    qImposto.Add("  JOIN SFN_IMPOSTO_ITEM ITEM ON ITEM.IMPOSTO = PI.IMPOSTO ")
		    	qImposto.Add(" WHERE (:DATA >= PI.DATAINICIAL AND (PI.DATAFINAL IS NULL OR :DATA <= PI.DATAFINAL)) ")
		    	qImposto.Add("   AND PI.PRESTADOR = :HANDLEPRESTADOR ")
		    	qImposto.Add("   AND ITEM.ITEM = :HANDLEITEMIMPOSTO ")
				qImposto.ParamByName("HANDLEPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
   		 		qImposto.ParamByName("HANDLEITEMIMPOSTO").AsInteger = CurrentQuery.FieldByName("ITEMNFRPAINTERNACAO").AsInteger
   		 		qImposto.ParamByName("DATA").Value = ServerDate
   	 			qImposto.Active = True

		    	If (qImposto.FieldByName("EXITEIMPOSTOVIGENTE").AsInteger = 0) Then
		    	    ValidarRegraEmissaoAutomaticaRPA = False
					mensagemRetorno = "Para emissão automatica de RPA escolher um item de internação de imposto vigente do prestador."
					Set qImposto = Nothing
					Exit Function
				End If

				Set qImposto = Nothing
			End If
		End If

		If (CurrentQuery.FieldByName("EMITENFPOR").AsInteger = 3) Then
			ValidarRegraEmissaoAutomaticaRPA = False
			mensagemRetorno = "A emissão do RPA é somente para quem emite nota fiscal."
			Exit Function
		End If

	End If

	ValidarRegraEmissaoAutomaticaRPA = True

End Function

Public Function ValidarItensNFParaRPA(vItemNF As Long) As Boolean
	Dim qImpostosOrigemAmbos As BPesquisa
    Set qImpostosOrigemAmbos = NewQuery

	qImpostosOrigemAmbos.Add("SELECT HANDLE                                                                 ")
	qImpostosOrigemAmbos.Add("  FROM SAM_PRESTADOR_IMPOSTO                                                  ")
	qImpostosOrigemAmbos.Add(" WHERE PRESTADOR = :HANDLEPRESTADOR                                           ")
	qImpostosOrigemAmbos.Add("   AND ORIGEMRECURSO = 'A'                                                    ")
	qImpostosOrigemAmbos.Add("   AND ( DATAINICIAL <= :DATA AND (DATAFINAL IS NULL OR DATAFINAL >= :DATA) ) ")

	qImpostosOrigemAmbos.ParamByName("HANDLEPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 	qImpostosOrigemAmbos.ParamByName("DATA").Value = ServerDate

	qImpostosOrigemAmbos.Active = True

	If (qImpostosOrigemAmbos.FieldByName("HANDLE").IsNull) Then
		Dim qImposto As BPesquisa
		Set qImposto = NewQuery

		qImposto.Add("SELECT X.HANDLE EXISTEIMPOSTOVIGENTE                                                   ")
		qImposto.Add("FROM (                                                                                 ")
		qImposto.Add("SELECT I.HANDLE                                                                        ")
		qImposto.Add("  FROM SAM_PRESTADOR_IMPOSTO PI                                                        ")
		qImposto.Add("  JOIN SFN_IMPOSTO           IM ON IM.HANDLE = PI.IMPOSTO                              ")
		qImposto.Add("  JOIN SFN_IMPOSTO_ITEM      II ON IM.HANDLE = II.IMPOSTO                              ")
		qImposto.Add("  JOIN SFN_ITEMNOTAFISCAL     I ON I.HANDLE = II.ITEM                                  ")
		qImposto.Add(" WHERE PI.PRESTADOR = :HANDLEPRESTADOR                                                 ")
		qImposto.Add("   AND ( PI.DATAINICIAL <= :DATA AND (PI.DATAFINAL IS NULL OR PI.DATAFINAL >= :DATA) ) ")
		qImposto.Add("   AND PI.ORIGEMRECURSO = 'P'                                                          ")
		qImposto.Add("   AND I.HANDLE = :HANDLEITEMIMPOSTO                                                   ")
		qImposto.Add("INTERSECT                                                                              ")
		qImposto.Add("SELECT I.HANDLE                                                                        ")
		qImposto.Add("  FROM SAM_PRESTADOR_IMPOSTO PI                                                        ")
		qImposto.Add("  JOIN SFN_IMPOSTO           IM ON IM.HANDLE = PI.IMPOSTO                              ")
		qImposto.Add("  JOIN SFN_IMPOSTO_ITEM      II ON IM.HANDLE = II.IMPOSTO                              ")
		qImposto.Add("  JOIN SFN_ITEMNOTAFISCAL     I ON I.HANDLE = II.ITEM                                  ")
		qImposto.Add(" WHERE PI.PRESTADOR = :HANDLEPRESTADOR                                                 ")
		qImposto.Add("   AND ( PI.DATAINICIAL <= :DATA AND (PI.DATAFINAL IS NULL OR PI.DATAFINAL >= :DATA) ) ")
		qImposto.Add("   AND PI.ORIGEMRECURSO = 'O'                                                          ")
		qImposto.Add("   AND I.HANDLE = :HANDLEITEMIMPOSTO                                                   ")
		qImposto.Add(") X                                                                                    ")

		qImposto.ParamByName("HANDLEPRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 		qImposto.ParamByName("HANDLEITEMIMPOSTO").AsInteger = vItemNF
	 	qImposto.ParamByName("DATA").Value = ServerDate

		qImposto.Active = True

    	If (qImposto.FieldByName("EXISTEIMPOSTOVIGENTE").IsNull) Then

            ValidarItensNFParaRPA = False

			qImposto.Active = False
			Set qImposto = Nothing

			qImpostosOrigemAmbos.Active = False
			Set qImpostosOrigemAmbos = Nothing
			Exit Function
		End If

		qImposto.Active = False
		Set qImposto = Nothing
	End If

	qImpostosOrigemAmbos.Active = False
	Set qImpostosOrigemAmbos = Nothing

	ValidarItensNFParaRPA = True

End Function

Public Function VerificaCPFDuplicado(handleCategoriaPrestador As Integer, cpfPrestador As String, tipoValidacao As String) As Boolean

	Dim Interface As CSEntityCall

	If tipoValidacao = "titular" Then
	    Set Interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfBeneficiarioTitularIgualCpfInformado")

	ElseIf tipoValidacao = "dependente" Then
		Set Interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfBeneficiarioDependenteIgualCpfInformado")

	ElseIf tipoValidacao = "usuario" Then
		Set Interface = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamCategoriaPrestador, Benner.Saude.Entidades", "VerificarCpfUsuarioIgualCpfInformado")

	End If

	Interface.AddParameter(pdtInteger, handleCategoriaPrestador)
	Interface.AddParameter(pdtString , cpfPrestador)

	VerificaCPFDuplicado = CBool(Interface.Execute())

	Set Interface = Nothing

End Function

Public Sub MostraAfastamento
	Dim AFAST As Object
	Set AFAST = NewQuery

	AFAST.Add("SELECT B.DESCRICAO, A.DATAFINAL")
	AFAST.Add("  FROM SAM_PRESTADOR_AFASTAMENTO A, SAM_MOTIVOAFASTAMENTO B")
	AFAST.Add(" WHERE A.PRESTADOR = :PRESTADOR")
	AFAST.Add("   AND B.HANDLE = A.MOTIVOAFASTAMENTO AND :DATA >= A.DATAINICIAL AND (:DATA <= A.DATAFINAL OR DATAFINAL IS NULL)")
	AFAST.Add(" ORDER BY A.DATAINICIAL")

	AFAST.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	AFAST.ParamByName("DATA").Value = ServerDate
	AFAST.Active = True

	LABELAFASTAMENTO.Text = ""

	If Not AFAST.EOF Then
		AFAST.First

		If AFAST.FieldByName("DATAFINAL").IsNull Then
			LABELAFASTAMENTO.Text = "Afastado - " + AFAST.FieldByName("DESCRICAO").AsString + " Indefinidamente"
		Else
			LABELAFASTAMENTO.Text = "Afastado - " + AFAST.FieldByName("DESCRICAO").AsString + " Até " + AFAST.FieldByName("DATAFINAL").AsString
		End If
	End If

	Set AFAST = Nothing
End Sub

Public Sub TABLE_NewRecord()

	Dim vFilial As Long
	Dim vFilialProc As Long
	Dim vMsg As String

	BuscarFiliais(CurrentSystem, vFilial, vFilialProc, vMsg)

	CurrentQuery.FieldByName("FILIALPADRAO").Value = vFilial
End Sub

'===============================IMPORTAÃ#O DE PROPONENTES =================================
Public Sub BOTAOIMPORTARPROPONENTE_OnClick()
  ImportarProponente

  Dim qParametrosPortal As BPesquisa
  Dim qParametrosPrestador As BPesquisa
  Dim qProponente As BPesquisa

  Set qParametrosPortal = NewQuery
  Set qProponente = NewQuery
  Set qParametrosPrestador = NewQuery

  qParametrosPortal.Clear
  qParametrosPortal.Active = False
  qParametrosPortal.Add(" SELECT EMAILINDICADOR ")
  qParametrosPortal.Add("	  FROM POR_CONFIGPORTAL")
  qParametrosPortal.Active = True

  qProponente.Active = False
  qProponente.Clear
  qProponente.Add(" SELECT * ")
  qProponente.Add("   FROM SAM_PROPONENTE ")
  qProponente.Add("  WHERE HANDLE = :HANDLE")
  qProponente.ParamByName("HANDLE").AsInteger = vgHandleProponente
  qProponente.Active = True

  qParametrosPrestador.Active = False
  qParametrosPrestador.Clear
  qParametrosPrestador.Add(" SELECT * ")
  qParametrosPrestador.Add("   FROM SAM_PARAMETROSPRESTADOR ")
  qParametrosPrestador.Active = True

  Dim dll As CSBusinessComponent
  Dim MensagemErro As String

  If (qParametrosPortal.FieldByName("EMAILINDICADOR").AsString = "S" And qProponente.FieldByName("EMAILINDICACAO").AsString <> "") Then
  	If (qParametrosPrestador.FieldByName("INICIOCREDENCIAMENTO").AsInteger <> 0) Then
  		Set dll = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.Proponente.EnvioEmailAoIndicadorDoProponente, Benner.Saude.Prestadores.Business")
  		dll.AddParameter(pdtInteger,vgHandleProponente)
 		dll.AddParameter(pdtInteger, qParametrosPrestador.FieldByName("INICIOCREDENCIAMENTO").AsInteger)
 		dll.AddParameter(pdtString,"Iniciado processo de credenciamento para o Proponente Indicado")
 		MensagemErro = dll.Execute("EnviarEmailAoIndicador")
	Else
	   bsShowMessage("O email não foi enviado ao indicador do proponente pois faltou parametrizar a mensagem à ser enviada. Para o indicador ser notificado, enviar email manualmente.","I")
	End If
  End If

  Set qParametrosPrestador = Nothing
  Set qParametrosPortal = Nothing
  Set qProponente = Nothing

  If (Len(MensagemErro) > 0) Then
	bsShowMessage(MensagemErro, "I")
  End If

Set dll = Nothing

End Sub

Public Sub ImportarProponente
  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterios As String
  Dim vCategoria As Long
  Dim vTpPrestador As Long
  Dim vNComplexidade As Long
  Dim vISSJuridica As Long
  Dim vISSFisica As Long
  Dim SQLPARAM As Object
  Set SQLPARAM = NewQuery
  Dim Msg As String
  Dim qSQLPROPONENTE As Object
  Dim DLLEspecifico As Object
  Dim vResultado As Boolean

  vgInsertProponente = False

  If VisibleMode Then
	If(checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N")Then
		bsShowMessage(Msg, "I")
		Exit Sub
	End If



	If CurrentQuery.State <>3 Then
		bsShowMessage("Para importar, primeiramente tecle NOVO registro.", "I")
		Exit Sub
	End If

	Set Interface = CreateBennerObject("Procura.Procurar")

	vColunas = "SAM_PROPONENTE.NOME|SAM_PROPONENTE.PROPONENTE"
	vCriterios = "SAM_PROPONENTE.SITUACAO='D' and SAM_PROPONENTE.PRESTADOR IS NULL  and sam_proponente.filial = Z_GRUPOUSUARIOS_FILIAIS.filial"
	vCampos = "Nome|CNPJ/CPF"
	vgHandleProponente = Interface.Exec(CurrentSystem, "SAM_PROPONENTE|Z_GRUPOUSUARIOS_FILIAIS[Z_GRUPOUSUARIOS_FILIAIS.USUARIO = SAM_PROPONENTE.INCLUSAOPOR]", vColunas, 1, vCampos, vCriterios, "Tabela de Proponentes", True, "")


	' INICIO SMS CONVENIOS Workflow - LEANDRO MANSO - VERIFICA PROPONETE PARA IMPORTAR
	Set DLLEspecifico = CreateBennerObject("ESPECIFICO.UESPECIFICO")

	vResultado = DLLEspecifico.PRE_ProponentePossuiEspecialidade(CurrentSystem, "SAM_PRESTADOR.BOTAOIMPORTARPROPONENTE", vgHandleProponente)

	If vResultado = False Then
  		bsShowMessage("Só é possível importar proponente que tenha ao menos uma especialidade principal que possa ser Importada no Prestador!", "E")
		'CanContinue = False
		Set DLLEspecifico = Nothing
		Exit Sub
	End If

	Set DLLEspecifico = Nothing
	' FIM SMS CONVENIOS Workflow
  End If

  SQLPARAM.Add("SELECT PROPONENTECATEGORIA, PROPONENTENIVELCOMPLEXIDADE, PROPONENTEISSFISICA, PROPONENTEISSJURIDICA FROM SAM_PARAMETROSPRESTADOR")

  SQLPARAM.Active = True

  If vgHandleProponente <>0 Then
	Set qSQLPROPONENTE = NewQuery

	qSQLPROPONENTE.Add("SELECT * FROM SAM_PROPONENTE WHERE HANDLE = :HANDLE")

	qSQLPROPONENTE.ParamByName("HANDLE").Value = vgHandleProponente
	qSQLPROPONENTE.Active = True

	If Not qSQLPROPONENTE.EOF Then
	  PrestadorRecebeProponenteOP("CPFCNPJ", "PROPONENTE", qSQLPROPONENTE, "S")

   	  CurrentQuery.FieldByName("NOME").Value = qSQLPROPONENTE.FieldByName("NOME").AsString
	  CurrentQuery.FieldByName("Z_NOME").Value = TiraAcento(qSQLPROPONENTE.FieldByName("NOME").AsString, True)
	  CurrentQuery.FieldByName("FANTASIA").Value = qSQLPROPONENTE.FieldByName("FANTASIA").AsString
	  CurrentQuery.FieldByName("Z_FANTASIA").Value = TiraAcento(qSQLPROPONENTE.FieldByName("FANTASIA").AsString, True)
	  CurrentQuery.FieldByName("FILIALPADRAO").Value = qSQLPROPONENTE.FieldByName("FILIAL").AsString
	  CurrentQuery.FieldByName("INCLUSAODATA").Value = ServerNow
	  CurrentQuery.FieldByName("DATAINCLUSAO").Value = ServerDate

  	  PrestadorRecebeProponenteOP("COOPERATIVA", "COOPERATIVA", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("CORPOCLINICO", "CORPOCLINICO", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("ASSOCIACAO", "ASSOCIACAO", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("INCLUSAOPOR", "INCLUSAOPOR", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("DATANASCIMENTO", "DATANASCIMENTO", qSQLPROPONENTE, "D")
	  PrestadorRecebeProponenteOP("ESTADOCIVIL", "ESTADOCIVIL", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("SOLICITANTE", "SOLICITANTE", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("RECEBEDOR", "RECEBEDOR", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("EXECUTOR", "EXECUTOR", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("FISICAJURIDICA", "FISICAJURIDICA", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("HOMEPAGE", "HOMEPAGE", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("EMAIL", "EMAIL", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("INSCRICAOESTADUAL", "INSCRICAOESTADUAL", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("INSCRICAOINSS", "INSCRICAOINSS", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("INSCRICAOMUNICIPAL", "INSCRICAOMUNICIPAL", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("ESTADO", "ESTADO", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("ESTADOPAGAMENTO", "ESTADO", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("MUNICIPIOPAGAMENTO", "MUNICIPIO", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("NACIONALIDADE", "NACIONALIDADE", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("NATURALIDADE", "NATURALIDADE", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("ORGAOEMISSOR", "ORGAOEMISSOR", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("CENTRALPAGER", "CENTRALPAGER", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("PAGER", "PAGER", qSQLPROPONENTE, "S")
      PrestadorRecebeProponenteOP("CELULAR", "CELULAR", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("RG", "RG", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("SEXO", "SEXO", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("TIPOPRESTADOR", "TIPOPRESTADOR", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("CONSELHOREGIONAL", "CONSELHOREGIONAL", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("UFCR", "UFCR", qSQLPROPONENTE, "I")
	  PrestadorRecebeProponenteOP("REGIAOCR", "REGIAOCR", qSQLPROPONENTE, "S")
	  PrestadorRecebeProponenteOP("INSCRICAOCR", "INSCRICAOCR", qSQLPROPONENTE, "S")
      PrestadorRecebeProponenteOP("DATAINSCRICAOCR", "DATAINSCRICAOCR", qSQLPROPONENTE, "D")

  	  If SQLPARAM.FieldByName("PROPONENTECATEGORIA").IsNull Then
		CurrentQuery.FieldByName("PROPONENTECATEGORIA").Clear
	  Else
		CurrentQuery.FieldByName("CATEGORIA").Value = SQLPARAM.FieldByName("PROPONENTECATEGORIA").AsInteger
	  End If

  	  If SQLPARAM.FieldByName("PROPONENTENIVELCOMPLEXIDADE").IsNull Then
		CurrentQuery.FieldByName("NIVELCOMPLEXIDADE").Clear
	  Else
		CurrentQuery.FieldByName("NIVELCOMPLEXIDADE").Value = SQLPARAM.FieldByName("PROPONENTENIVELCOMPLEXIDADE").AsInteger
	  End If

  	  If qSQLPROPONENTE.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
		CurrentQuery.FieldByName("ISS").Value = SQLPARAM.FieldByName("PROPONENTEISSFISICA").AsInteger
	  Else
		CurrentQuery.FieldByName("ISS").Value = SQLPARAM.FieldByName("PROPONENTEISSJURIDICA").AsInteger
	  End If

  	  vgInsertProponente = True
	End If
  End If
End Sub

Public Function PrestadorRecebeProponenteOP(pCampoPRE, pCampoPRO As String, pQuery As Object, pTipo As String)As Boolean
	If pQuery.FieldByName(pCampoPRO).IsNull Then
		CurrentQuery.FieldByName(pCampoPRE).Clear
	Else
		Select Case pTipo
			Case "S"
				CurrentQuery.FieldByName(pCampoPRE).Value = pQuery.FieldByName(pCampoPRO).AsString
			Case "I"
				CurrentQuery.FieldByName(pCampoPRE).Value = pQuery.FieldByName(pCampoPRO).AsInteger
			Case "D"
				CurrentQuery.FieldByName(pCampoPRE).Value = pQuery.FieldByName(pCampoPRO).AsDateTime
		End Select
	End If
End Function

Public Sub INSERTENDERECO()
	Dim SQLEND As Object
	Dim qSQLPROPONENTE As BPesquisa
	Set qSQLPROPONENTE = NewQuery
	Set SQLEND = NewQuery

	If vgHandleProponente <>0 Then
		Set qSQLPROPONENTE = NewQuery

		qSQLPROPONENTE.Clear
		qSQLPROPONENTE.Add("SELECT * FROM SAM_PROPONENTE WHERE HANDLE = :HANDLE")

		qSQLPROPONENTE.ParamByName("HANDLE").Value = vgHandleProponente
		qSQLPROPONENTE.Active = True

		SQLEND.Clear

		SQLEND.Add("INSERT INTO SAM_PRESTADOR_ENDERECO (HANDLE, PRESTADOR, CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,")
		SQLEND.Add("			NUMERO, COMPLEMENTO, PONTOREFERENCIA,TELEFONE1,FAX, CORRESPONDENCIA, ATENDIMENTO,QTDVAGASESTACIONAMENTO, CNES)")
		SQLEND.Add("VALUES (:HANDLE, :PRESTADOR, :CEP, :ESTADO, :MUNICIPIO, :BAIRRO, :LOGRADOURO, :NUMERO,")
		SQLEND.Add("        :COMPLEMENTO, :PONTOREFERENCIA, :TELEFONE1, :FAX, :CORRESPONDENCIA, :ATENDIMENTO,:QTDVAGASESTACIONAMENTO, :CNES)")

		vgHandleEnderecoProponente = NewHandle("SAM_PRESTADOR_ENDERECO")

		SQLEND.ParamByName("HANDLE").Value = vgHandleEnderecoProponente
		SQLEND.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
		SQLEND.ParamByName("CEP").Value = qSQLPROPONENTE.FieldByName("CEP").AsString
		SQLEND.ParamByName("CNES").Value = qSQLPROPONENTE.FieldByName("CNES").AsString
		SQLEND.ParamByName("ESTADO").Value = qSQLPROPONENTE.FieldByName("ESTADO").AsInteger
		SQLEND.ParamByName("MUNICIPIO").Value = qSQLPROPONENTE.FieldByName("MUNICIPIO").AsInteger
		SQLEND.ParamByName("BAIRRO").Value = qSQLPROPONENTE.FieldByName("BAIRRO").AsString
		SQLEND.ParamByName("LOGRADOURO").Value = qSQLPROPONENTE.FieldByName("LOGRADOURO").AsString
		SQLEND.ParamByName("NUMERO").Value = qSQLPROPONENTE.FieldByName("LOGRADOURONUMERO").AsInteger
		SQLEND.ParamByName("COMPLEMENTO").Value = qSQLPROPONENTE.FieldByName("LOGRADOUROCOMPLEMENTO").AsString
		SQLEND.ParamByName("PONTOREFERENCIA").Value = qSQLPROPONENTE.FieldByName("PONTOREFERENCIA").AsString
		SQLEND.ParamByName("CORRESPONDENCIA").Value = qSQLPROPONENTE.FieldByName("CORRESPONDENCIA").AsString
		SQLEND.ParamByName("ATENDIMENTO").Value = qSQLPROPONENTE.FieldByName("ATENDIMENTO").AsString
		SQLEND.ParamByName("TELEFONE1").Value = qSQLPROPONENTE.FieldByName("TELEFONE").AsString
		SQLEND.ParamByName("FAX").Value = qSQLPROPONENTE.FieldByName("FAX").AsString
		SQLEND.ParamByName("QTDVAGASESTACIONAMENTO").Value = 0

		SQLEND.ExecSQL

	End If

	Set SQLEND = Nothing
End Sub

Public Sub INSERTCURRICULO()
	Dim SQLFORM As Object
	Dim SQLCUR As Object
	Set SQLFORM = NewQuery

	'FORMACAO Do PROPONENTE
	SQLFORM.Add("SELECT * FROM SAM_PROPONENTE_FORMACAO WHERE PROPONENTE = :HANDLE")

	SQLFORM.ParamByName("HANDLE").Value = vgHandleProponente
	SQLFORM.Active = True

	While Not(SQLFORM.EOF)
		'CURRICULO Do PRESTADOR
		Set SQLCUR = NewQuery

		SQLCUR.Clear

		SQLCUR.Add("INSERT INTO SAM_PRESTADOR_CURRICULO (HANDLE, PRESTADOR, DATAINICIAL,DATACONCLUSAO, CARGAHORARIA,TIPOCURSO, AREACURSO,CURSO,ENTIDADE,OBSERVACAO )")
		SQLCUR.Add("VALUES (:HANDLE, :PRESTADOR, :DATAINICIAL,:DATACONCLUSAO, :CARGAHORARIA, :TIPOCURSO, :AREACURSO,:CURSO,:ENTIDADE,:OBSERVACAO)")

		SQLCUR.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_CURRICULO")
		SQLCUR.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
		SQLCUR.ParamByName("DATAINICIAL").Value = SQLFORM.FieldByName("DATAINICIAL").AsDateTime

	    If Not SQLFORM.FieldByName("DATACONCLUSAO").IsNull Then
    		SQLCUR.ParamByName("DATACONCLUSAO").Value = SQLFORM.FieldByName("DATACONCLUSAO").AsDateTime  ' SMS 91312 - Paulo Melo - 11/01/2008
    	Else																							 ' incluindo prestador através de proponente
	    	SQLCUR.ParamByName("DATACONCLUSAO").DataType = ftDate										 ' o campo data de conclusão no curriculo
    		SQLCUR.ParamByName("DATACONCLUSAO").Clear												     ' ficava com a data default de 12/1899
    	End If																							 ' agora fica em branco.

		SQLCUR.ParamByName("CARGAHORARIA").Value = SQLFORM.FieldByName("CARGAHORARIA").AsInteger
		SQLCUR.ParamByName("TIPOCURSO").Value = SQLFORM.FieldByName("TIPOCURSO").AsString
		SQLCUR.ParamByName("AREACURSO").Value = SQLFORM.FieldByName("AREACURSO").AsInteger
		SQLCUR.ParamByName("CURSO").Value = SQLFORM.FieldByName("CURSO").AsString
		SQLCUR.ParamByName("ENTIDADE").Value = SQLFORM.FieldByName("ENTIDADE").AsString
		SQLCUR.ParamByName("OBSERVACAO").Value = SQLFORM.FieldByName("OBSERVACAO").AsString

		SQLCUR.ExecSQL

		SQLCUR.Active = False

		SQLFORM.Next
	Wend

	Set SQLCUR = Nothing
	Set SQLFORM = Nothing
End Sub

Public Sub INSERTCURRICULOEXPERIENCIA()
	Dim SQLFORM As Object
	Dim SQLCUREMPRES As Object
	Set SQLFORM = NewQuery

	'FORMACAO Do PROPONENTE
	SQLFORM.Add("SELECT * FROM SAM_PROPONENTE_EXPERIENCIA WHERE PROPONENTE = :HANDLE")

	SQLFORM.ParamByName("HANDLE").Value = vgHandleProponente
	SQLFORM.Active = True

	While Not(SQLFORM.EOF)
		'CURRICULO Do PRESTADOR
		Set SQLCUREMPRES = NewQuery

		SQLCUREMPRES.Clear

		SQLCUREMPRES.Add("INSERT INTO SAM_PRESTADOR_CURRICULO_EXP (HANDLE, PRESTADOR,NOME,Z_NOME,ESTADO,MUNICIPIO,BAIRRO, ")
		SQLCUREMPRES.Add("LOGRADOURO,LOGRADOURONUMERO,LOGRADOUROCOMPLEMENTO,PONTOREFERENCIA,TELEFONE,CEP,DATAINICIAL,DATAFINAL)")
		SQLCUREMPRES.Add("VALUES (:HANDLE, :PRESTADOR,:NOME,:Z_NOME,:ESTADO,:MUNICIPIO,:BAIRRO, ")
		SQLCUREMPRES.Add(":LOGRADOURO,:LORADOURONUMERO,:LOGRADOUROCOMPLEMENTO,:PONTOREFERENCIA,:TELEFONE,:CEP,:DATAINICIAL,:DATAFINAL)")

		SQLCUREMPRES.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_CURRICULO_EXP")
		SQLCUREMPRES.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
		SQLCUREMPRES.ParamByName("NOME").Value = SQLFORM.FieldByName("NOME").AsString
		SQLCUREMPRES.ParamByName("Z_NOME").Value = SQLFORM.FieldByName("NOME").AsString
		SQLCUREMPRES.ParamByName("ESTADO").Value = SQLFORM.FieldByName("ESTADO").AsInteger
		SQLCUREMPRES.ParamByName("MUNICIPIO").Value = SQLFORM.FieldByName("MUNICIPIO").AsInteger
		SQLCUREMPRES.ParamByName("BAIRRO").Value = SQLFORM.FieldByName("BAIRRO").AsString
		SQLCUREMPRES.ParamByName("LOGRADOURO").Value = SQLFORM.FieldByName("LOGRADOURO").AsString
		SQLCUREMPRES.ParamByName("LORADOURONUMERO").Value = SQLFORM.FieldByName("LORADOURONUMERO").AsInteger
		SQLCUREMPRES.ParamByName("LOGRADOUROCOMPLEMENTO").Value = SQLFORM.FieldByName("LOGRADOUROCOMPLEMENTO").AsString
		SQLCUREMPRES.ParamByName("PONTOREFERENCIA").Value = SQLFORM.FieldByName("PONTOREFERENCIA").AsString
		SQLCUREMPRES.ParamByName("TELEFONE").Value = SQLFORM.FieldByName("TELEFONE").AsString
		SQLCUREMPRES.ParamByName("CEP").Value = SQLFORM.FieldByName("CEP").AsString
		SQLCUREMPRES.ParamByName("DATAINICIAL").Value = SQLFORM.FieldByName("DATAINICIAL").AsDateTime

		If Not SQLFORM.FieldByName("DATAFINAL").IsNull Then
    		SQLCUREMPRES.ParamByName("DATAFINAL").Value = SQLFORM.FieldByName("DATAFINAL").AsDateTime  ' SMS 91312 - Paulo Melo - 11/01/2008
    	Else																					 			' incluindo prestador através de proponente
	    	SQLCUREMPRES.ParamByName("DATAFINAL").DataType = ftDate									        ' o campo data final do curriculo experiência
    		SQLCUREMPRES.ParamByName("DATAFINAL").Clear												 		' ficava com a data default de 12/1899
    	End If																					 			' agora fica em branco.

		SQLCUREMPRES.ExecSQL

		SQLCUREMPRES.Active = False

		SQLFORM.Next
	Wend

	Set SQLCUREMPRES = Nothing
	Set SQLFORM = Nothing
End Sub

Public Sub UPDATEPROPONENTE()
	Dim qSQL As Object
	Set qSQL = NewQuery

		Dim qParametrosPortal As BPesquisa
  	Dim qParametrosPrestador As BPesquisa
  	Dim qProponente As BPesquisa

	Set qParametrosPortal = NewQuery
    Set qProponente = NewQuery
  	Set qParametrosPrestador = NewQuery

  	qParametrosPortal.Clear
  	qParametrosPortal.Active = False
  	qParametrosPortal.Add(" SELECT EMAILINDICADOR ")
  	qParametrosPortal.Add("	  FROM POR_CONFIGPORTAL")
  	qParametrosPortal.Active = True

  	qProponente.Active = False
  	qProponente.Clear
  	qProponente.Add(" SELECT * ")
  	qProponente.Add("   FROM SAM_PROPONENTE ")
  	qProponente.Add("  WHERE HANDLE = :HANDLE")
  	qProponente.ParamByName("HANDLE").AsInteger = vgHandleProponente
  	qProponente.Active = True

  	qParametrosPrestador.Active = False
  	qParametrosPrestador.Clear
  	qParametrosPrestador.Add(" SELECT * ")
  	qParametrosPrestador.Add("   FROM SAM_PARAMETROSPRESTADOR ")
  	qParametrosPrestador.Active = True

   	Dim dll As CSBusinessComponent
	Dim MensagemErro As String

	If (qParametrosPortal.FieldByName("EMAILINDICADOR").AsString = "S" And qProponente.FieldByName("EMAILINDICACAO").AsString <> "") Then
  		If (qParametrosPrestador.FieldByName("CREDENCIAMENTODEFERIDO").AsInteger <> 0) Then
  			Set dll = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.Proponente.EnvioEmailAoIndicadorDoProponente, Benner.Saude.Prestadores.Business")
  			dll.AddParameter(pdtInteger,vgHandleProponente)
  			dll.AddParameter(pdtInteger, qParametrosPrestador.FieldByName("CREDENCIAMENTODEFERIDO").AsInteger)
  			dll.AddParameter(pdtString,"Proponente Indicado foi credenciado")
  			MensagemErro = dll.Execute("EnviarEmailAoIndicador")
		Else
	   	    bsShowMessage("O email não foi enviado ao indicador do proponente pois faltou parametrizar a mensagem à ser enviada. Para o indicador ser notificado, enviar email manualmente.","I")
	    End If
    End If

  	Set qParametrosPrestador = Nothing
  	Set qParametrosPortal = Nothing
  	Set qProponente = Nothing

	If (Len(MensagemErro) > 0) Then
		bsShowMessage(MensagemErro, "I")
  	End If

	Set dll = Nothing

	'altera numero Do PRESTADOR,significa que proponente virou PRESTADOR
	qSQL.Clear

	qSQL.Add("UPDATE SAM_PROPONENTE SET PRESTADOR = :HANDLEPRESTADOR, SITUACAO = 'C', DATACADASTRO = :DATACADASTRO WHERE HANDLE= :HANDLEPROPONENTE ")

	qSQL.ParamByName("HANDLEPROPONENTE").Value = vgHandleProponente
	qSQL.ParamByName("HANDLEPRESTADOR").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	qSQL.ParamByName("DATACADASTRO").Value = ServerDate

	qSQL.ExecSQL

	Set qSQL = Nothing
End Sub

Public Sub INSERTESPECIALIZADES
	Dim vHandlePrestadorEspecialidade As Long
	Dim vHandlePrestadorEspecialidadeGrupo As Long
	Dim vHandlePrestadorEspecialidadeGrupoReg As Long
	Dim vHandlePrestadorEspecialidadeGrupoRede As Long
	Dim vHandlePrestadorEspecialidadeRede As Long
	Dim vHandlePrestadorLivro As Long
	Dim qEspecialidade As Object
	Set qEspecialidade = NewQuery
	Dim qEspecialidadeGrupo As Object
	Set qEspecialidadeGrupo = NewQuery
	Dim qEspecialidadeGrpReg As Object
	Set qEspecialidadeGrpReg = NewQuery
	Dim qEspecialidadeGrpRede As Object
	Set qEspecialidadeGrpRede = NewQuery
	Dim qEspecialidadeRede As Object
	Set qEspecialidadeRede = NewQuery
	Dim qInsert As Object
	Set qInsert = NewQuery

	If CurrentQuery.FieldByName ("DATAINCLUSAO").IsNull Then
		bsShowMessage("Data de inclusão é obrigatória, por favor verifique!", "E")
		Exit Sub
	End If

	On Error GoTo Erro

	qEspecialidade.Active = False

	qEspecialidade.Clear

	qEspecialidade.Add ("SELECT A.*")
	qEspecialidade.Add ("  FROM SAM_PROPONENTE_ESPECIALIDADE A")
	qEspecialidade.Add (" WHERE A.PROPONENTE = :PROPONENTE")
	qEspecialidade.Add ("   AND A.TABIMPORTAR = 1")

	qEspecialidade.ParamByName ("PROPONENTE").AsInteger = vgHandleProponente
	qEspecialidade.Active = True

	While Not(qEspecialidade.EOF)
		qInsert.Active = False

		qInsert.Clear

		qInsert.Add ("INSERT INTO SAM_PRESTADOR_ESPECIALIDADE (HANDLE")
		qInsert.Add ("                                        ,Z_GRUPO")
		qInsert.Add ("                                        ,ESPECIALIDADE")
		qInsert.Add ("                                        ,PRESTADOR")
		qInsert.Add ("                                        ,PRINCIPAL")
		qInsert.Add ("                                        ,TEMPORARIO")
		qInsert.Add ("                                        ,DATAINICIAL")
		qInsert.Add ("                                        ,PUBLICARNOLIVRO")
		qInsert.Add ("                                        ,PUBLICARINTERNET")
		qInsert.Add ("                                        ,VISUALIZARCENTRAL)")
		qInsert.Add ("VALUES                                  (:HANDLE")
		qInsert.Add ("                                        ,:Z_GRUPO")
		qInsert.Add ("                                        ,:ESPECIALIDADE")
		qInsert.Add ("                                        ,:PRESTADOR")
		qInsert.Add ("                                        ,:PRINCIPAL")
		qInsert.Add ("                                        ,:TEMPORARIO")
		qInsert.Add ("                                        ,:DATAINICIAL")
		qInsert.Add ("                                        ,:PUBLICARNOLIVRO")
		qInsert.Add ("                                        ,:PUBLICARINTERNET")
		qInsert.Add ("                                        ,:VISUALIZARCENTRAL)")

		vHandlePrestadorEspecialidade = NewHandle("SAM_PRESTADOR_ESPECIALIDADE")

		qInsert.ParamByName("HANDLE").AsInteger = vHandlePrestadorEspecialidade
		qInsert.ParamByName("Z_GRUPO").AsInteger = qEspecialidade.FieldByName("Z_GRUPO").AsInteger
		qInsert.ParamByName("ESPECIALIDADE").AsInteger = qEspecialidade.FieldByName("ESPECIALIDADE").AsInteger
		qInsert.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		qInsert.ParamByName("PRINCIPAL").AsString = qEspecialidade.FieldByName("PRINCIPAL").AsString
		qInsert.ParamByName("TEMPORARIO").AsString = "N"
		qInsert.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime
		qInsert.ParamByName("PUBLICARNOLIVRO").AsString = qEspecialidade.FieldByName("PUBLICARNOLIVRO").AsString
		qInsert.ParamByName("PUBLICARINTERNET").AsString = qEspecialidade.FieldByName("PUBLICARINTERNET").AsString
		qInsert.ParamByName("VISUALIZARCENTRAL").AsString = qEspecialidade.FieldByName("VISUALIZARCENTRAL").AsString

		qInsert.ExecSQL

		If qEspecialidade.FieldByName("PUBLICARNOLIVRO").AsString = "S" Then
			qInsert.Active = False

			qInsert.Clear

			qInsert.Add ("INSERT INTO SAM_PRESTADOR_LIVRO (HANDLE")
			qInsert.Add ("                                ,PRESTADOR")
			qInsert.Add ("                                ,AREA")
			qInsert.Add ("                                ,ESPECIALIDADE")
			qInsert.Add ("                                ,ENDERECO")
			qInsert.Add ("                                ,PUBLICARNOLIVRO")
			qInsert.Add ("                                ,PUBLICARINTERNET")
			qInsert.Add ("                                ,VISUALIZARCENTRAL)")
			qInsert.Add ("VALUES                          (:HANDLE")
			qInsert.Add ("                                ,:PRESTADOR")
			qInsert.Add ("                                ,:AREA")
			qInsert.Add ("                                ,:ESPECIALIDADE")
			qInsert.Add ("                                ,:ENDERECO")
			qInsert.Add ("                                ,:PUBLICARNOLIVRO")
			qInsert.Add ("                                ,:PUBLICARINTERNET")
			qInsert.Add ("                                ,:VISUALIZARCENTRAL)")

			vHandlePrestadorLivro = NewHandle("SAM_PRESTADOR_LIVRO")

			qInsert.ParamByName("HANDLE").AsInteger = vHandlePrestadorLivro
			qInsert.ParamByName("ESPECIALIDADE").AsInteger = qEspecialidade.FieldByName("ESPECIALIDADE").AsInteger
			qInsert.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			qInsert.ParamByName("AREA").AsInteger = qEspecialidade.FieldByName("AREALIVRO").AsInteger
			qInsert.ParamByName("ENDERECO").AsInteger = vgHandleEnderecoProponente
			qInsert.ParamByName("PUBLICARNOLIVRO").AsString = "S"
			qInsert.ParamByName("PUBLICARINTERNET").AsString = "S"
			qInsert.ParamByName("VISUALIZARCENTRAL").AsString = "S"

			qInsert.ExecSQL
		End If

		qEspecialidadeGrupo.Active = False
		qEspecialidadeGrupo.Clear
		qEspecialidadeGrupo.Add ("SELECT A.*")
		qEspecialidadeGrupo.Add ("  FROM SAM_PROPONENTE_ESPEC_GRP A")
		qEspecialidadeGrupo.Add (" WHERE A.PROPONENTEESPEC = :PROPONENTEESPEC")
		qEspecialidadeGrupo.ParamByName ("PROPONENTEESPEC").AsInteger = qEspecialidade.FieldByName ("HANDLE").AsInteger
		qEspecialidadeGrupo.Active = True

		While Not (qEspecialidadeGrupo.EOF)
			qInsert.Active = False

			qInsert.Clear

			qInsert.Add ("INSERT INTO SAM_PRESTADOR_ESPECIALIDADEGRP (HANDLE")
			qInsert.Add ("                                           ,Z_GRUPO")
			qInsert.Add ("                                           ,PRESTADOR")
			qInsert.Add ("                                           ,ESPECIALIDADEGRUPO")
			qInsert.Add ("                                           ,ESPECIALIDADE")
			qInsert.Add ("                                           ,PRESTADORESPECIALIDADE")
			qInsert.Add ("                                           ,DATAINICIAL")
			qInsert.Add ("                                           ,PERMITERECEBER")
			qInsert.Add ("                                           ,PERMITEEXECUTAR)")
			qInsert.Add ("VALUES                                     (:HANDLE")
			qInsert.Add ("                                           ,:Z_GRUPO")
			qInsert.Add ("                                           ,:PRESTADOR")
			qInsert.Add ("                                           ,:ESPECIALIDADEGRUPO")
			qInsert.Add ("                                           ,:ESPECIALIDADE")
			qInsert.Add ("                                           ,:PRESTADORESPECIALIDADE")
			qInsert.Add ("                                           ,:DATAINICIAL")
			qInsert.Add ("                                           ,:PERMITERECEBER")
			qInsert.Add ("                                           ,:PERMITEEXECUTAR)")

			vHandlePrestadorEspecialidadeGrupo = NewHandle("SAM_PRESTADOR_ESPECIALIDADEGRP")

			qInsert.ParamByName("HANDLE").AsInteger = vHandlePrestadorEspecialidadeGrupo
			qInsert.ParamByName("Z_GRUPO").AsInteger = qEspecialidadeGrupo.FieldByName("Z_GRUPO").AsInteger
			qInsert.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			qInsert.ParamByName("ESPECIALIDADE").AsInteger = qEspecialidadeGrupo.FieldByName("ESPECIALIDADE").AsInteger
			qInsert.ParamByName("ESPECIALIDADEGRUPO").AsInteger = qEspecialidadeGrupo.FieldByName("ESPECIALIDADEGRUPO").AsInteger
			qInsert.ParamByName("PRESTADORESPECIALIDADE").AsInteger = vHandlePrestadorEspecialidade
			qInsert.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime
			qInsert.ParamByName("PERMITERECEBER").AsString = qEspecialidadeGrupo.FieldByName("PERMITERECEBER").AsString
			qInsert.ParamByName("PERMITEEXECUTAR").AsString = qEspecialidadeGrupo.FieldByName("PERMITEEXECUTAR").AsString

			qInsert.ExecSQL

			qEspecialidadeGrpReg.Active = False

			qEspecialidadeGrpReg.Clear

			qEspecialidadeGrpReg.Add ("SELECT A.*")
			qEspecialidadeGrpReg.Add ("  FROM SAM_PROPONENTE_ESPEC_GRP_REG A")
			qEspecialidadeGrpReg.Add (" WHERE A.PROPONENTEESPECGRP = :PROPONENTEESPECGRP")

			qEspecialidadeGrpReg.ParamByName ("PROPONENTEESPECGRP").AsInteger = qEspecialidadeGrupo.FieldByName ("HANDLE").AsInteger
			qEspecialidadeGrpReg.Active = True

			While Not (qEspecialidadeGrpReg.EOF)
				qInsert.Active = False

				qInsert.Clear

				qInsert.Add ("INSERT INTO SAM_PRESTADOR_ESPECIALIDADEREG (HANDLE")
				qInsert.Add ("                                           ,Z_GRUPO")
				qInsert.Add ("                                           ,PRESTADORESPECIALIDADEGRP")
				qInsert.Add ("                                           ,REGIMEATENDIMENTO)")
				qInsert.Add ("VALUES                                     (:HANDLE")
				qInsert.Add ("                                           ,:Z_GRUPO")
				qInsert.Add ("                                           ,:PRESTADORESPECIALIDADEGRP")
				qInsert.Add ("                                           ,:REGIMEATENDIMENTO)")

				vHandlePrestadorEspecialidadeGrupoReg = NewHandle("SAM_PRESTADOR_ESPECIALIDADEREG")

				qInsert.ParamByName("HANDLE").AsInteger = vHandlePrestadorEspecialidadeGrupoReg
				qInsert.ParamByName("Z_GRUPO").AsInteger = qEspecialidadeGrpReg.FieldByName("Z_GRUPO").AsInteger
				qInsert.ParamByName("PRESTADORESPECIALIDADEGRP").AsInteger = vHandlePrestadorEspecialidadeGrupo
				qInsert.ParamByName("REGIMEATENDIMENTO").AsInteger = qEspecialidadeGrpReg.FieldByName("REGIMEATENDIMENTO").AsInteger

				qInsert.ExecSQL

				qEspecialidadeGrpReg.Next
			Wend

			qEspecialidadeGrpRede.Active = False

			qEspecialidadeGrpRede.Clear

			qEspecialidadeGrpRede.Add ("SELECT A.*")
			qEspecialidadeGrpRede.Add ("  FROM SAM_PROPONENTE_ESPEC_GRP_REDE A")
			qEspecialidadeGrpRede.Add (" WHERE A.PROPONENTEESPECGRP = :PROPONENTEESPECGRP")

			qEspecialidadeGrpRede.ParamByName ("PROPONENTEESPECGRP").AsInteger = qEspecialidadeGrupo.FieldByName ("HANDLE").AsInteger

			qEspecialidadeGrpRede.Active = True

			While Not (qEspecialidadeGrpRede.EOF)
				qInsert.Active = False

				qInsert.Clear

				qInsert.Add ("INSERT INTO SAM_PRESTADOR_ESPEC_GRP_REDE   (HANDLE")
				qInsert.Add ("                                           ,Z_GRUPO")
				qInsert.Add ("                                           ,PRESTADORESPECIALIDADEGRUPO")
				qInsert.Add ("                                           ,REDE)")
				qInsert.Add ("VALUES                                     (:HANDLE")
				qInsert.Add ("                                           ,:Z_GRUPO")
				qInsert.Add ("                                           ,:PRESTADORESPECIALIDADEGRUPO")
				qInsert.Add ("                                           ,:REDE)")

				vHandlePrestadorEspecialidadeGrupoRede = NewHandle("SAM_PRESTADOR_ESPEC_GRP_REDE")

				qInsert.ParamByName("HANDLE").AsInteger = vHandlePrestadorEspecialidadeGrupoRede
				qInsert.ParamByName("Z_GRUPO").AsInteger = qEspecialidadeGrpRede.FieldByName("Z_GRUPO").AsInteger
				qInsert.ParamByName("PRESTADORESPECIALIDADEGRUPO").AsInteger = vHandlePrestadorEspecialidadeGrupo
				qInsert.ParamByName("REDE").AsInteger = qEspecialidadeGrpRede.FieldByName("REDERESTRITA").AsInteger

				qInsert.ExecSQL

				qEspecialidadeGrpRede.Next
			Wend

			qEspecialidadeGrupo.Next
		Wend

		qEspecialidadeRede.Active = False

		qEspecialidadeRede.Clear

		qEspecialidadeRede.Add ("SELECT A.*")
		qEspecialidadeRede.Add ("  FROM SAM_PROPONENTE_ESPEC_REDE A")
		qEspecialidadeRede.Add (" WHERE A.PROPONENTEESPEC = :PROPONENTEESPEC")

		qEspecialidadeRede.ParamByName ("PROPONENTEESPEC").AsInteger = qEspecialidade.FieldByName ("HANDLE").AsInteger

		qEspecialidadeRede.Active = True

		While Not (qEspecialidadeRede.EOF)
			qInsert.Active = False

			qInsert.Clear

			qInsert.Add ("INSERT INTO SAM_PRESTADOR_ESPEC_REDE (HANDLE")
			qInsert.Add ("                                     ,Z_GRUPO")
			qInsert.Add ("                                     ,PRESTADORESPECIALIDADE")
			qInsert.Add ("                                     ,REDERESTRITA")
			qInsert.Add ("                                     ,DATAINICIAL)")
			qInsert.Add ("VALUES                               (:HANDLE")
			qInsert.Add ("                                     ,:Z_GRUPO")
			qInsert.Add ("                                     ,:PRESTADORESPECIALIDADE")
			qInsert.Add ("                                     ,:REDERESTRITA")
			qInsert.Add ("                                     ,:DATAINICIAL)")

			vHandlePrestadorEspecialidadeRede = NewHandle("SAM_PRESTADOR_ESPEC_REDE")

			qInsert.ParamByName("HANDLE").AsInteger = vHandlePrestadorEspecialidadeRede
			qInsert.ParamByName("Z_GRUPO").AsInteger = qEspecialidadeRede.FieldByName("Z_GRUPO").AsInteger
			qInsert.ParamByName("PRESTADORESPECIALIDADE").AsInteger = vHandlePrestadorEspecialidade
			qInsert.ParamByName("REDERESTRITA").AsInteger = qEspecialidadeRede.FieldByName("REDERESTRITA").AsInteger
			qInsert.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime

			qInsert.ExecSQL

			qEspecialidadeRede.Next
		Wend

		qEspecialidade.Next
	Wend

	Erro:

	Fim:
		Set qEspecialidade = Nothing
		Set qEspecialidadeGrupo = Nothing
		Set qEspecialidadeGrpReg = Nothing
		Set qEspecialidadeGrpRede = Nothing
		Set qEspecialidadeRede = Nothing
		Set qInsert = Nothing
End Sub

Public Sub ISS_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If CurrentQuery.State <>1 Then CurrentQuery.UpdateRecord
		If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 0 Then
			bsShowMessage("Favor escolher antes o tipo do prestador (Física ou Jurídica)", "I")
			ShowPopup = False
			Exit Sub
		End If

	If ISS.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "ISS"
		vColunas = "DESCRICAO"

		If Trim(CurrentQuery.FieldByName("FISICAJURIDICA").AsString)<>"" Then
			vCriterio = "(FISICAJURIDICA = " + CurrentQuery.FieldByName("FISICAJURIDICA").AsString + ")"
		Else
			vCriterio = ""
		End If

		vTabela = "SFN_ISS"
		vTitulo = "Tipo de ISS"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("ISS").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub CATEGORIA_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If CATEGORIA.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Código|Categorias|Considera dim. vagas"
		vColunas = "CODIGO|DESCRICAO|CONSIDERADIMVAGAS"
		vCriterio = ""
		vTabela = "SAM_CATEGORIA_PRESTADOR"
		vTitulo = "Categoria do prestador"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("CATEGORIA").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub MOTIVOBLOQUEIO_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If MOTIVOBLOQUEIO.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Motivo"
		vColunas = "DESCRICAO"
		vCriterio = ""
		vTabela = "SAM_MOTIVOBLOQUEIO"
		vTitulo = "Motivo de bloqueio"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("MOTIVOBLOQUEIO").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub MOTIVOREFERENCIAMENTO_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If MOTIVOREFERENCIAMENTO.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Motivo|Peso"
		vColunas = "DESCRICAO|PESO"
		vCriterio = ""
		vTabela = "SAM_MOTIVOREFERENCIAMENTO"
		vTitulo = "Motivo de referenciamento"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("MOTIVOREFERENCIAMENTO").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub NIVELCOMPLEXIDADE_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If NIVELCOMPLEXIDADE.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Nível de complexidade"
		vColunas = "DESCRICAO"
		vCriterio = ""
		vTabela = "SAM_NIVELCOMPLEXIDADE"
		vTitulo = "Nível de complexidade"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("NIVELCOMPLEXIDADE").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCRIARUSUARIOWEB"
			BOTAOCRIARUSUARIOWEB_OnClick

		Case "BOTAOINICIARCREDENCIAMENTO"
			Dim vMsg As String
			vMsg = ValidarPermissoesBotaoIniciarCredenciamento(CurrentQuery)

			If (vMsg <> "") Then
				bsshowmessage( vMsg, "E")
				CanContinue = False
			End If
	End Select
End Sub

Public Sub TABLE_UpdateRequired()
  If CurrentQuery.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
    CurrentQuery.FieldByName("CPFCNPJ").Mask = "999\.999\.999\-99;0;_"
  Else
	CurrentQuery.FieldByName("CPFCNPJ").Mask = "99\.999\.999\/9999\-99;0;_"
  End If

  'Remover a formatação do campo CPF/CNPJ na inserção de registros em modo web
  If WebMode And _
     CurrentQuery.State = 3 _
     And Not CurrentQuery.FieldByName("CPFCNPJ").IsNull Then
    CurrentQuery.FieldByName("CPFCNPJ").AsString = Replace(Replace(Replace(CurrentQuery.FieldByName("CPFCNPJ").AsString, ".", ""), "/", ""), "-", "")
  End If
End Sub

Public Sub TIPOPRESTADOR_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If TIPOPRESTADOR.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Código|Tipo|Exige parecer do representante|Exige parecer da sede"
		vColunas = "SAM_TIPOPRESTADOR.CODIGO|SAM_TIPOPRESTADOR.DESCRICAO|SAM_TIPOPRESTADOR.EXIGEPARECERREPRESENTACAO|SAM_TIPOPRESTADOR.EXIGEPARECERSEDE"
		vColunas = vColunas
		vCriterio = ""
		vTabela = "SAM_TIPOPRESTADOR"
		vTitulo = "Tipo do prestador"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("TIPOPRESTADOR").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub CONVERSAOABRAMGE_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If CONVERSAOABRAMGE.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Código|Descrição"
		vColunas = "CODIGO|DESCRICAO"
		vCriterio = ""
		vTabela = "SAM_CONVERSAO"
		vTitulo = "Conversão"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("CONVERSAOABRAMGE").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub CONSELHOREGIONAL_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If CONSELHOREGIONAL.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Sigla|Descrição"
		vColunas = "SIGLA|DESCRICAO"
		vCriterio = ""
		vTabela = "SAM_CONSELHO"
		vTitulo = "Conselho regional"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("CONSELHOREGIONAL").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub MUNICIPIOPAGAMENTO_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If MUNICIPIOPAGAMENTO.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Cidade"
		vColunas = "MUNICIPIOS.NOME"

		If Trim(CurrentQuery.FieldByName("ESTADOPAGAMENTO").AsString)<>"" Then
			vCriterio = "(MUNICIPIOS.ESTADO = " + CurrentQuery.FieldByName("ESTADOPAGAMENTO").AsString + ")"
		Else
			vCriterio = ""
		End If

		vTabela = "MUNICIPIOS"
		vTitulo = "Municipios"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("MUNICIPIOPAGAMENTO").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub USUARIOCONFERENCIAPEG_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If USUARIOCONFERENCIAPEG.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Código|Nome|Apelido"
		vColunas = "CODIGO|NOME|APELIDO"
		vCriterio = ""
		vTabela = "Z_GRUPOUSUARIOS"
		vTitulo = "Usuários"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("USUARIOCONFERENCIAPEG").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub PRESTADORMESTRE_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If PRESTADORMESTRE.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Código|Prestador|CPFCNPJ"
		vColunas = "PRESTADOR|NOME|CPFCNPJ"
		vCriterio = ""
		vTabela = "SAM_PRESTADOR"
		vTitulo = "Grupo empresarial"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCabecs, vCriterio, vTitulo, False, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("PRESTADORMESTRE").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub NATURALIDADE_OnPopup(ShowPopup As Boolean)
	Dim Interface As Object
	Dim vHandle As Long
	Dim vCabecs As String
	Dim vColunas As String
	Dim vCriterio As String
	Dim vTabela As String
	Dim vTitulo As String

	If NATURALIDADE.PopupCase <>0 Then
		Set Interface = CreateBennerObject("Procura.Procurar")

		ShowPopup = False
		vCabecs = "Naturalidade|UF|País"
		vColunas = "MUNICIPIOS.NOME|ESTADOS.SIGLA|PAISES.NOME"

		If Trim(CurrentQuery.FieldByName("ESTADO").AsString)<>"" Then
			vCriterio = "(MUNICIPIOS.ESTADO = " + CurrentQuery.FieldByName("ESTADO").AsString + ")"
		Else
			vCriterio = ""
		End If

		vTabela = "MUNICIPIOS|ESTADOS[MUNICIPIOS.ESTADO = ESTADOS.HANDLE]|PAISES[PAISES.HANDLE=MUNICIPIOS.PAIS]"
		vTitulo = "Grupo empresarial"
		vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCabecs, vCriterio, vTitulo, True, "")

		If vHandle <>0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("NATURALIDADE").AsInteger = vHandle
		End If

		Set Interface = Nothing
	Else
		ShowPopup = True
	End If
End Sub

Public Sub BOTAOINICIARCREDENCIAMENTO_OnClick()
   	If CurrentQuery.State = 2 Then
    	bsShowMessage("O registro está em edição, salve-o para acessar esta funcionalidade", "E")
		Exit Sub
	End If

	Dim vMsg As String
	vMsg = ValidarPermissoesBotaoIniciarCredenciamento(CurrentQuery)

	If (vMsg <> "") Then
		bsshowmessage( vMsg, "E")
		Exit Sub
	End If

	Dim vTipoProcesso As Long
	Dim vInseriuProcesso As Long
	vInseriuProcesso = -1

	On Error GoTo Except

		vMsg = ""
		vTipoProcesso = BuscarTiposDeProcessoDeCredenciamento(CurrentQuery)
	    If (vTipoProcesso <> -1) Then

	    	vInseriuProcesso = InserirProcessoCredenciamentoInicial(CurrentQuery.FieldByName("HANDLE").AsInteger ,vTipoProcesso, ServerDate, 0, "S")
	    Else
			Dim Interface As Object
			Dim viRetorno As Integer
	    	Dim vsMensagem As String
	    	Dim vvContainer As CSDContainer
	    	Set vvContainer = NewContainer
	    	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
	    	viRetorno = Interface.Exec(CurrentSystem, _
	                             1, _
	                             "TV_CREDENCIAR_PRESTADOR", _
	                             "Credenciar Prestador", _
	                             0, _
	                             250, _
	                             380, _
	                             False, _
	                             vsMensagem, _
	                             vvContainer)
	    	Set Interface = Nothing

			If (viRetorno <> -1) Then

	    		vInseriuProcesso = InserirProcessoCredenciamentoInicial(CurrentQuery.FieldByName("HANDLE").AsInteger, vvContainer.Field("TIPOCREDENCIAMENTO").AsInteger, _
		  			vvContainer.Field("DATACREDENCIAMENTO").AsDateTime, vvContainer.Field("NOVAFILIAL").AsInteger, vvContainer.Field("INCLUIRFASES").AsString)
			End If
	    End If
	    GoTo Fim
	Except:
		vInseriuProcesso = -2
		vMsg = Err.Description
	Fim:
	On Error GoTo 0

	Select Case vInseriuProcesso
		Case 0
			bsShowMessage("Inclusão de Credenciamento Concluída!", "I")
			If VisibleMode Then
				RefreshNodesWithTable("SAM_PRESTADOR")
			End If
		Case -1

			bsShowMessage("Processo cancelado!", "I")

		Case -2

			vMsg = "Falha ao inserir processo de credenciamento: " + Chr(13) + vMsg
			bsShowMessage(vMsg, "E")
	End Select
End Sub
