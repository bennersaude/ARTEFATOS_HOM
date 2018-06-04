'HASH: 0B4703E9D01F65C80E95C34148BFADAF
'#Uses "*bsShowMessage

Option Explicit

Public Sub CONFERENCIANF_OnClick()
 	Dim dllValidarMensagem As Object
	Set dllValidarMensagem = CreateBennerObject("Benner.Saude.ProcessamentoContas.Formularios.Start")
	dllValidarMensagem.ShowConferenciaNF
	Set dllValidarMensagem = Nothing
End Sub

Public Sub CONSULTABENEFICIARIO_OnClick()
	Dim interface As Object
	'Alterado SMS 90338 - Rodrigo Andrade 30/11/2007 -
	'Separação da Interface da regra de negocio para consulta de Beneficiários
	Set interface =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
	interface.filtro(CurrentSystem,1,"")
End Sub

Public Sub CONSULTAEVENTOPACOTE_OnClick()
	Dim OLESamObject As Object
	Set OLESamObject =CreateBennerObject("SamConsulta.Consulta")
	OLESamObject.EventosPacote(CurrentSystem,-1,-1,1950 -1 -1)
	Set OLESamObject =Nothing
End Sub

Public Sub CONSULTAPRECOEVENTO_OnClick()
  Dim interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String
  Dim vvContainer As CSDContainer

  Set interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
  viRetorno = interface.Exec(CurrentSystem, _
                             1, _
                             "TV_FORM0004", _
                             "Consulta Preço do Evento", _
                             0, _
                             740, _
                             640, _
                             False, _
                             vsMensagem, _
                             vvContainer)

  Set interface =Nothing
End Sub

Public Sub CONSULTAPRESTADORES_OnClick()
	Dim interface As Object
	Dim vsMensagem As String
	Dim vlHPrestador As Long
	Set interface =CreateBennerObject("BSINTERFACE0001.BuscaPrestador")

	interface.Abrir(CurrentSystem, vsMensagem, 1, "", "T", vlHPrestador)
End Sub

Public Sub CONSULTAPRESTADOREVE_OnClick()
    Dim QueryBuscaHandledoRelatorio As Object
	Dim HandleDoRelatorio As Long
	Set QueryBuscaHandledoRelatorio =NewQuery

	QueryBuscaHandledoRelatorio.Add("SELECT R.HANDLE RELATORIO_HANDLE FROM R_RELATORIOS R WHERE R.CODIGO = 'PRE001'")
	QueryBuscaHandledoRelatorio.Active =False
	QueryBuscaHandledoRelatorio.Active =True
	HandleDoRelatorio =QueryBuscaHandledoRelatorio.FieldByName("RELATORIO_HANDLE").AsInteger
	Set QueryBuscaHandledoRelatorio =Nothing

	ReportPreview(HandleDoRelatorio,"",False,False)
End Sub

Public Sub CONSULTAREGRAPREST_OnClick()
	Dim QueryBuscaHandledoRelatorio As Object
	Dim rep As CSReportPrinter
	Set QueryBuscaHandledoRelatorio =NewQuery

	QueryBuscaHandledoRelatorio.Add("SELECT R.HANDLE RELATORIO_HANDLE FROM R_RELATORIOS R WHERE R.CODIGO = 'PRE006X'")
	QueryBuscaHandledoRelatorio.Active =False
	QueryBuscaHandledoRelatorio.Active =True

	Set rep = NewReport(QueryBuscaHandledoRelatorio.FieldByName("RELATORIO_HANDLE").AsInteger)

	rep.Preview

    Set QueryBuscaHandledoRelatorio =Nothing
	Set rep = Nothing

End Sub

Public Sub DIGITABENEFICIARIO_OnClick()
  'Crislei.sorrilha SMS 108068
  Dim dllBSInterface001 As Object

  Set dllBSInterface001 = CreateBennerObject("BSINTERFACE0011.DigitarBeneficiario")

  dllBSInterface001.Exec(CurrentSystem, _
						 0, _
						 0, _
						 0)

  Set dllBSInterface001 = Nothing
End Sub

Public Sub EXPORTARECOM_OnClick()
	'SMS 81133 - Ricardo Rocha - 16/08/2007
	Dim SAMECOM As Object
	Set SAMECOM = CreateBennerObject("SAMECOM.Rotinas")
	SAMECOM.Exec(CurrentSystem)
	Set SAMECOM = Nothing
End Sub

Public Sub EXTRATPREVINNE_OnClick()
  'Luciano T. Alberti - SMS 103070 - 29/09/2008 - Início
  Dim interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String
  Dim vvContainer As CSDContainer

  Set interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
  viRetorno = interface.Exec(CurrentSystem, _
                             1, _
                             "TV_EXTRPREVINE", _
                             "Gerar arquivo ao PREVINNE", _
                             0, _
                             200, _
                             340, _
                             False, _
                             vsMensagem, _
                             vvContainer)

  Set interface =Nothing
  'Luciano T. Alberti - SMS 103070 - 29/09/2008 - Fim
End Sub

Public Sub LIBERARPAGAMENTO_OnClick()

 Dim LiberarPagto As CSBusinessComponent

 Set LiberarPagto = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.Pagamentos.SamAgrupadorPagamentoBLL, Benner.Saude.ProcessamentoContas.Business")
 LiberarPagto.Execute("LiberarPagamento")

End Sub

Public Sub LIBERATETOALCADA_OnClick()
  Dim interface As Object

  Set interface = CreateBennerObject("LiberacaoTetoAlcadaPagamento.LiberacaoTetoAlcadaPagamento")
  interface.Exec(0, ServerDate)
  Set interface = Nothing
End Sub

Public Sub MODULE_BeforeNodeShow(ByVal NodeFullPath As String, CanShow As Boolean)

	If (NodeFullPath = "4.6 Faturamento|PAGAMENTO") Then
		Dim qVerificaParametroImpostoNF As Object
		Set qVerificaParametroImpostoNF = NewQuery
		qVerificaParametroImpostoNF.Add("SELECT TABIMPOSTOSNANF,          ")
		qVerificaParametroImpostoNF.Add("       TABCONCILIACAODOCFISCAIS, ")
		qVerificaParametroImpostoNF.Add("       TABCONTROLEPAGAMENTO      ")
		qVerificaParametroImpostoNF.Add("  FROM SAM_PARAMETROSPROCCONTAS  ")
		qVerificaParametroImpostoNF.Active = True

		If (qVerificaParametroImpostoNF.FieldByName("TABIMPOSTOSNANF").AsInteger = 1) And _
		   (qVerificaParametroImpostoNF.FieldByName("TABCONCILIACAODOCFISCAIS").AsInteger = 1) And _
		   (qVerificaParametroImpostoNF.FieldByName("TABCONTROLEPAGAMENTO").AsInteger = 1) Then
		  CanShow = True
		Else
		  CanShow = False
		End If

		Set qVerificaParametroImpostoNF = Nothing
	End If

	If InStr(NodeFullPath, "TABELAS|MOTIVOS_DE_GLOSA|EXCEPCIONALIDADE|MOTIVOS_GLOSA") > 0 Then
		CanShow = False
   	End If

	If (NodeFullPath = "COMPETENCIASDEPAGAMENTO|FILIAIS|6.9.1.1-CARGAPAGAMENTOSSITUACOES|BLOQUEADOS|COMPOSICAO" Or _
		NodeFullPath = "COMPETENCIASDEPAGAMENTO|FILIAIS|6.9.1.1-CARGAPAGAMENTOSSITUACOES|LIBERADOS|COMPOSICAO" Or _
		NodeFullPath = "COMPETENCIASDEPAGAMENTO|FILIAIS|6.9.1.1-CARGAPAGAMENTOSSITUACOES|AGUARDANDOFATURAMENTO|COMPOSICAO" Or _
		NodeFullPath = "COMPETENCIASDEPAGAMENTO|FILIAIS|6.9.1.1-CARGAPAGAMENTOSSITUACOES|FATURADOS|COMPOSICAO" Or _
		NodeFullPath = "COMPETENCIASDEPAGAMENTO|FILIAIS|6.9.1.1-CARGAPAGAMENTOSSITUACOES|TODOS|COMPOSICAO") Then
		Dim qVerificaParametroDotacaoOrcamentaria As Object
		Set qVerificaParametroDotacaoOrcamentaria = NewQuery
		qVerificaParametroDotacaoOrcamentaria.Add("SELECT CONTROLADOTORC     ")
		qVerificaParametroDotacaoOrcamentaria.Add("  FROM SFN_PARAMETROSFIN  ")
		qVerificaParametroDotacaoOrcamentaria.Active = True

		If (qVerificaParametroDotacaoOrcamentaria.FieldByName("CONTROLADOTORC").AsInteger = 2) Then
		  CanShow = True
		Else
		  CanShow = False
		End If

		Set qVerificaParametroDotacaoOrcamentaria = Nothing
	End If

End Sub

Public Sub MODULE_OnEnter()

	MONITORCONTAS.Visible = False
    MONITORPAGAMENTO.Visible = UtilizaDotacaoOrcamentaria

	Dim vDllBSPro006 As Object

  	Set vDllBSPro006 = CreateBennerObject("BSPro006.Rotinas")

  	DIGITACAOPEG.Visible = vDllBSPro006.PermitirDigitacaoPeg(CurrentSystem)

  	Set vDllBSPro006 = Nothing

End Sub

Public Sub MONITORCONTAS_OnClick()
	Dim DLL As Object
    Set DLL = CreateBennerObject("BSINTERFACE0064.CONFERENCIA")

    DLL.Monitor

    Set DLL = Nothing
End Sub

Public Sub MONITORPAGAMENTO_OnClick()

 Dim MonitorPagto As CSBusinessComponent

 Set MonitorPagto = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.Pagamentos.SamMonitorPagamentoBLL, Benner.Saude.ProcessamentoContas.Business")
 MonitorPagto.Execute("MonitorPagamento")

 Set MonitorPagto = Nothing
End Sub

Public Function UtilizaDotacaoOrcamentaria As Boolean
	Dim qDotacaoOrcamentaria As BPesquisa
	Set qDotacaoOrcamentaria = NewQuery

	UtilizaDotacaoOrcamentaria = False

	qDotacaoOrcamentaria.Add("SELECT CONTROLADOTORC FROM SFN_PARAMETROSFIN")
	qDotacaoOrcamentaria.Active = True

	If (qDotacaoOrcamentaria.FieldByName("CONTROLADOTORC").AsInteger = 2) Then
	  UtilizaDotacaoOrcamentaria = True
	End If

	Set qDotacaoOrcamentaria = Nothing

End Function

Public Sub PRESTLIVREESCOLHA_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_FORM0066", _
								   "Prestador Livre-escolha", _
								   0, _
								   480, _
								   640, _
								   False, _
								   vsMensagem, _
								   vcContainer)

  Set vcContainer = Nothing
End Sub


Public Sub PRESTINTERCAMBIO_OnClick()
	Dim interface As Object
	Set interface = CreateBennerObject("SAMPRESTLIVREESCOLHA.CADASTRAPRESTADOR")
	interface.CadastroManual_PrestIntercambio(CurrentSystem,0)
	Set interface =  Nothing
End Sub


Public Sub PROCESSOPEGLOTE_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("BSINTERFACE0039.ProcessarPegLote")
  interface.Exec(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub RECLASSIFICAR_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SamPEG.PROCESSAR")
  interface.Reclassificacao(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub REAPRESENTACAOEVENTO_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("BSINTERFACE0050.REAPRESENTACAO")

  interface.Exec(CurrentSystem)

  Set interface = Nothing
End Sub

Public Sub DIGITACAOPEG_OnClick()

	Dim vlHandle As Long
	Dim vsTipo   As String

	vsTipo = "E"
	vlHandle = RecordHandleOfTable("SAM_GUIA_EVENTOS")

	If vlHandle <= 0 Then
    	vsTipo = "G"
		vlHandle = RecordHandleOfTable("SAM_GUIA")

	    If vlHandle <= 0 Then
			vsTipo = "P"
			vlHandle  = RecordHandleOfTable("SAM_PEG")

			If vlHandle <= 0 Then
				vsTipo = ""
			End If
		End If
	End If

  	Dim vDllBSPro006 As Object

  	Set vDllBSPro006 = CreateBennerObject("BSPro006.Rotinas")

  	vDllBSPro006.DigitacaoPeg(CurrentSystem, vlHandle, vsTipo)

  	Set vDllBSPro006 = Nothing

End Sub
