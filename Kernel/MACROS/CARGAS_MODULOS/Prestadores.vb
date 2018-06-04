'HASH: A1D17AAA46D9523C4F9D8E5FDA933CF4
 
 '#Uses "*bsShowMessage

Option Explicit

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


Public Sub CONSULTAPRECOMODULO_OnClick()
  Exit Sub
  'Dim INTERFACE As Object
  'Set INTERFACE =CreateBennerObject("RELATORIOS.RELS")
  'INTERFACE.EXEC(1,DBID)
  'Set INTERFACE=Nothing
End Sub


Public Sub CONSULTAPRESTADOR_OnClick()
	Dim interface As Object
	Dim vlHPrestador As Long
	Dim vsMensagem As String
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

	If(QueryBuscaHandledoRelatorio.FieldByName("RELATORIO_HANDLE").AsInteger > 0) Then
		Set rep = NewReport(QueryBuscaHandledoRelatorio.FieldByName("RELATORIO_HANDLE").AsInteger)
		rep.Preview
	Else
		bsShowMessage("Relatório não encontrado.", "I")
	End If

    Set QueryBuscaHandledoRelatorio =Nothing
	Set rep = Nothing

End Sub

Public Sub DETALHESPRESTADOR_OnClick()
Dim interface As Object
Set interface =CreateBennerObject("CA005.ConsultaPrestador")
interface.info(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)
End Sub

Public Sub GERARESPECIALIDADES_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SAMPROCPRESTADOR.PROCESSOPRESTADOR")
  interface.GerarEspecialidades(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub INCLUIORIGEMPREST_OnClick()

  Dim interface As Object
  Set interface=CreateBennerObject("BSPRE009.Rotinas")
  interface.OrigemEvento(CurrentSystem)
  Set interface=Nothing

End Sub

Public Sub PREFERENCIAMINIMA_OnClick()
  Dim interface As Object
  Set interface =CreateBennerObject("SamVaga.Atendimento")
  interface.PrefMin(CurrentSystem)
  Set interface =Nothing
End Sub

Public Sub MODULE_BeforeNodeShow(ByVal NodeFullPath As String, CanShow As Boolean)
	Dim qConsulta As BPesquisa

	Dim vCargaDocumentosEspecialidade As String
	Dim vCargaDocumentosCorpoClinico As String
	Dim vCargaDocumentosServContratado As String

	vCargaDocumentosEspecialidade = "3.5 Processos|CREDENCIAMENTO_E_ALTERACOES|TIPOS_DE_PROCESSO|CADASTR._DE_ESPECIALIDADE_OU_SERVICO|DOCUMENTOS_ENTREGUES"
	vCargaDocumentosCorpoClinico = "3.5 Processos|CREDENCIAMENTO_E_ALTERACOES|TIPOS_DE_PROCESSO|MEMBROS_DO_CORPO_CLINICO|DOCUMENTOSENTREGUESMEMBRO"
    vCargaDocumentosServContratado = "3.0_CADASTRO_DE_PRESTADORES|3_TODOS|3.5 Processos|CREDENCIAMENTO_E_ALTERACOES|TIPOS_DE_PROCESSO|SERVICOSCONTRATADOS|DOCUMENTOS_ENTREGUES_SERVCONT"

'	bsShowMessage("PATH " & NodeFullPath, "I")

	If Right(NodeFullPath, 29) = "|3.5.3.1.4_DOCUMENTOSEXIGIDOS" Then
		Dim vQueryParametros As BPesquisa
		Set vQueryParametros = NewQuery

		vQueryParametros.Add("SELECT P.CREDENCIAMENTOAVANCADO,                    ")
	    vQueryParametros.Add("COALESCE(T.CONTROLARDOCUMENTACAO,'N') CONTROLA_DOCS ")
	    vQueryParametros.Add("   FROM SAM_PARAMETROSPRESTADOR P,                  ")
	    vQueryParametros.Add("   SAM_TIPOPROCESSOCREDENCTO T                      ")
	    vQueryParametros.Add("WHERE T.HANDLE =                                    ")
	    vQueryParametros.Add("   (SELECT TIPOCREDENCIAMENTO                       ")
	    vQueryParametros.Add("      FROM SAM_PRESTADOR_PROC_CREDEN                ")
	    vQueryParametros.Add("   WHERE HANDLE = :HANDLEPROC)                      ")

     	vQueryParametros.ParamByName("HANDLEPROC").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN")
		vQueryParametros.Active = True
 		CanShow = vQueryParametros.FieldByName("CREDENCIAMENTOAVANCADO").AsInteger = 1 And vQueryParametros.FieldByName("CONTROLA_DOCS").AsString = "S"
		Set vQueryParametros = Nothing
	End If

	If Right(NodeFullPath, 29) = "|SAM_TIPOCREDENCIAMENTO_ESPEC" Then
		Dim vQueryRatificacao As BPesquisa
		Set vQueryRatificacao = NewQuery

		vQueryRatificacao.Add("SELECT CONTROLARFASERATIFICACAO,                                           ")
		vQueryRatificacao.Add("COALESCE(RATIFICARPORESPECIALIDADE, 'N') RATIFICARPORESPECIALIDADE         ")
		vQueryRatificacao.Add("       FROM SAM_TIPOCREDENCIAMENTO_FASE                                    ")
		vQueryRatificacao.Add("WHERE HANDLE = :PHANDLE                                                    ")

		vQueryRatificacao.ParamByName("PHANDLE").Value = RecordHandleOfTable("SAM_TIPOCREDENCIAMENTO_FASE")
		vQueryRatificacao.Active = True
		CanShow = vQueryRatificacao.FieldByName("CONTROLARFASERATIFICACAO").AsInteger = 1 And vQueryRatificacao.FieldByName("RATIFICARPORESPECIALIDADE").AsString = "S"
		Set vQueryRatificacao = Nothing
	End If

	If InStr(NodeFullPath, "3.0_CADASTRO_DE_PRESTADORES") > 0 And InStr(NodeFullPath, "EXCEPCIONALIDADE|PRESTADOR") > 0 Then
		CanShow = False
	End If

	If InStr(NodeFullPath, "3.2 Tabelas|TIPO_DE_PRESTADOR|EXCEPCIONALIDADE|TIPO_PRESTADOR") > 0 Then
		CanShow = False
	End If

	If InStr(NodeFullPath, vCargaDocumentosEspecialidade) > 0 Or InStr(NodeFullPath, vCargaDocumentosCorpoClinico) > 0 Or InStr(NodeFullPath, vCargaDocumentosServContratado) > 0 Then
		Dim qProcessoPortal As BPesquisa
		Set qProcessoPortal = NewQuery

		qProcessoPortal.Add("SELECT PROCESSOPORTAL                ")
		qProcessoPortal.Add("  FROM SAM_PRESTADOR_PROC            ")
		qProcessoPortal.Add(" WHERE HANDLE = :HANDLEPROCESSO      ")

		qProcessoPortal.ParamByName("HANDLEPROCESSO").AsInteger = RecordHandleOfTable("SAM_PRESTADOR_PROC")
		qProcessoPortal.Active = True

		CanShow = qProcessoPortal.FieldByName("PROCESSOPORTAL").AsString = "S"

		Set qProcessoPortal = Nothing
	End If


	If (InStr(NodeFullPath, "3.6.PA.4_PRESTADORESREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.4_PRESTADORESREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.4_PRESTADORESREAJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.4_PRESTADORESREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.4_PRESTADORESREAJUSTEPACOTE") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.1_ESTADOSREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.1_ESTADOSREJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.1_ESTADOSREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.1_ESTADOSREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.1_ESTADOSREAJUSTEPACOTE") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.3_ASSOCIACOESREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.3_ASSOCIACOESREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.3_ASSOCIACOESREJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.3_ASSOCIACOESREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.3_ASSOCIACOESREAJUSTEPACOTE") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.2_MUNICIPIOSREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.2_MUNICIPIOSREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.2_MUNICIPIOSREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.2_MUNICIPIOSREJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.2_MUNICIPIOSREAJUSTEPACOTE") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.5_REDEREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.5_REDEREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.5_REDEREAJUSTEREGIME") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.5_REDEREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.5_REDEREAJUSTEPACOTE") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.6_REDEPRESTADOR_REAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.6_REDEPRESTADOR_REAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.6_REDEPRESTADOR_REAJUSTEREG") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PA.6_REDEPRESTADOR_REAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "REAJUSTEDOTACAO|FAIXA_DE_EVENTOS_A_REAJUSTAR") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.4_PRESTADORESREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.4_PRESTADORESREAJUSTEREGATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.4_PRESTADORESREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.4_PRESTADORESREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.1_ESTADOSREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.1_ESTADOSREJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.1_ESTADOSREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.1_ESTADOSREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.3_ASSOCIACOESREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.3_ASSOCIACOESREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.3_ASSOCIACOESREJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.3_ASSOCIACOESREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.2_MUNICIPIOSREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.2_MUNICIPIOSREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.2_MUNICIPIOSREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.2_MUNICIPIOSREJUSTEREGIMEATEN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.5_REDEREAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.5_REDEREAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.5_REDEREAJUSTEREGIME") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.5_REDEREAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.6_REDEPRESTADOR_REAJUSTEAN") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.6_REDEPRESTADOR_REAJUSTEGRAU") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.6_REDEPRESTADOR_REAJUSTEREG") > 0) Or _
	   (InStr(NodeFullPath, "3.6.PF.6_REDEPRESTADOR_REAJUSTESL") > 0) Or _
	   (InStr(NodeFullPath, "REJUASTEDOTACAO|FAIXA_DE_EVENTOS_A_REAJUSTAR") > 0) Or _
	   (InStr(NodeFullPath, "REAJUSTEDOTAC|FAIXA_DE_EVENTOS_A_REAJUSTAR") > 0) Or _
	   (InStr(NodeFullPath, "3.6 - Preços|REAJUSTE_DE_PRECOS|3.6.Pa - Parâmetros (Abertos)|PRESTADORES|FILIADOS") > 0)Then

		Set qConsulta = NewQuery

		qConsulta.Active = False
		qConsulta.Clear
		qConsulta.Add("SELECT SRP.TIPOROTINA                     ")
		qConsulta.Add("  FROM SAM_REAJUSTEPRC_PARAM SRP          ")
		qConsulta.Add(" WHERE SRP.HANDLE = :PHANDLE              ")
		qConsulta.ParamByName("PHANDLE").AsInteger = RecordHandleOfTable("SAM_REAJUSTEPRC_PARAM")
		qConsulta.Active = True

		If (qConsulta.FieldByName("TIPOROTINA").AsInteger = 2) Then
    	   CanShow = False
    	End If

    	Set qConsulta = Nothing
    End If
End Sub
