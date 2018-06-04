'HASH: 1EE757510D7F1531802A9D73FDAF8A23
'MACRO='Macro: SAM_PEG
'sam_peg
'#USES "*ProcuraBeneficiarioAtivoReembolso"
'#Uses "*CheckCPFCNPJ"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraPrestador"
'#USES "*PrimeiroDiaCompetencia"
'#USES "*UltimoDiaCompetencia"
'#uses "*PermissaoAlteracao"
'#Uses "*TV_FORM0143_VALIDACAO"
'#Uses "*FuncoesDotacaoOrcamentaria"
'#Uses "*VerificarBloqueioAlteracoes"
'#Uses "*LimpaEspaco"
'#Uses "*VerificarBloqueioAlteracoesReapresentacao"
'#Uses "*VerificarBloqueioPegExisteComposicao"
'#Uses "*RecordHandleOfTableInterfacePEG"
'#Uses "*RefreshNodesWithTableInterfacePEG"
'#Uses "*SelectNodeInterfacePEG"
'#Uses "*VerificarPegReapresentadoExercicioPosterior"

Option Explicit

Dim Estadodatabela As Long
Dim OLDDATARECEBIMENTO As Date
Dim OLDRECEBEDOR As Long
Dim vOLDBeneficiarioTitular As Long
Dim OLDDATAPAGAMENTO As Date
Dim datautil As Date
Dim processachange As Boolean
Dim processachange2 As Boolean
Dim gchangeValorAdi As Boolean
Dim gchangeDataAdi As Boolean
Dim ACHOU As Boolean
Dim AUXNUMPAG As Long
Dim DATAx As Date
Dim EstavaAberta As Boolean
Dim vgFilial As Long
Dim vgFilialProcessamento As Long
Dim vgUtilizaCalendarioDiario As String
Dim vgCalendarioExcecao As String
Dim OldQtdGuia As Long
Dim TrocouNumeracaoPEG As Boolean
Dim ERROIDENTIFICADOR As Boolean
Dim OLDLOCALEXECUCAO As Long
Dim vbEditando As Boolean
Dim vgTemFilialCusto As Boolean
Dim qParametros As Object
Dim vVerificaAlteracoesCT As String
Dim qVerificaConsiderarSp As Object
Dim qServicosPrestador As Object
Dim qPreencheServicos As Object
Dim vgBanco As String
Dim vgAgencia As String
Dim vgContaCorrenteNumero As String
Dim vgContaCorrenteDV As String
Dim vgContaCorrenteNome As String
Dim vgContaCorrenteCPFCNPJ As String
Dim vTratarAlteracaoCredito As Boolean
Dim vPodeSAlvarPegDataPagamentoCorreta As Boolean
Dim viHTipoPegAnterior As Integer 'variavel para informar o tipo de peg anterior,  ñ alterar!
Dim vSituacaoAnteriorPeg As String
Dim vVerificaDataWeb As Boolean
Dim vTipoPegAnterior As Integer
Dim vBeneficiarioAnterior As Long
Dim VTabRegimePgtoAnterior As Integer
Dim vMudouTabRegimePagto As Boolean
Dim vMudouDataRecebimentoPEG As Boolean
Dim vMudouDataPagamentoPEG As Boolean
Dim bUtilizaTriagem As Boolean
Dim qSamParametrosProcContas As BPesquisa
Dim gTabOrigemRecursoPEG As Integer
Const cCabesp As Integer = 2


Public Function ProcuraBeneficiarioAtivo(pSoAtivos As Boolean,  pData As Date, TextoBenAtivo As String) As Long
  Dim Interface As Object
  Dim vWhere As String
  Dim vColunas As String
  Dim vDllDigit As Object
  Set qParametros=NewQuery
  Dim vOrdemBusca As Integer
  Set vDllDigit = CreateBennerObject("SAMPEGDIGIT.DIGITACAO")

  ProcuraBeneficiarioAtivo = vDllDigit.ValidarBeneficiario(CurrentSystem, TextoBenAtivo)

  If ProcuraBeneficiarioAtivo <= 0 Then
    qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
    qParametros.Active=True

    If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="S" Then
      'Set Interface=CreateBennerObject("CA010.ConsultaBeneficiario")
      'Separação da Interface da regra de negocio para consulta de Beneficiários
      Set Interface =CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")
      ProcuraBeneficiarioAtivo=Interface.Filtro(CurrentSystem,1,TextoBenAtivo)

      Set Interface=Nothing
    End If


    If qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString="N" Then
      vColunas = "SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_CONTRATO.CONTRATANTE|SAM_BENEFICIARIO.CODIGODEAFINIDADE|SAM_BENEFICIARIO.CODIGOANTIGO|SAM_BENEFICIARIO.DATACANCELAMENTO|SAM_CONVENIO.DESCRICAO|SAM_BENEFICIARIO.CODIGODEORIGEM|SAM_BENEFICIARIO.CODIGODEREPASSE"

      vWhere = ""

      If pSoAtivos = True Then
         vWhere = vWhere + "(SAM_BENEFICIARIO.DATABLOQUEIO IS NULL) And "
         vWhere = vWhere + " ((SAM_BENEFICIARIO.ATENDIMENTOATE Is NOT NULL AND SAM_BENEFICIARIO.ATENDIMENTOATE >= "+SQLDate(pData)+") OR (SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL OR SAM_BENEFICIARIO.DATACANCELAMENTO >= "+SQLDate(pData)+"))"
      End If

      vOrdemBusca = 2

      If IsNumeric(TextoBenAtivo) Then
        Dim qParamBenef As Object
        Dim qConsulta   As Object
        Dim vContador   As Integer
        Dim vPrioridade As String


        Set qParamBenef = NewQuery
        Set qConsulta   = NewQuery

        qParamBenef.Clear
        qParamBenef.Add("Select PRIORIDADE1, PRIORIDADE2, PRIORIDADE3, PRIORIDADE4, PRIORIDADE5")
        qParamBenef.Add("FROM SAM_PARAMETROSBENEFICIARIO")
        qParamBenef.Active = True

        vContador = 1

        While (vContador <= 5)
          vPrioridade = "PRIORIDADE" + Trim(Str(vContador))
          qConsulta.Clear
          qConsulta.Add("SELECT HANDLE FROM SAM_BENEFICIARIO")

          If qParamBenef.FieldByName(vPrioridade).AsInteger > 0 Then
            vOrdemBusca = 0

            Select Case qParamBenef.FieldByName(vPrioridade).AsInteger
              Case 1
                qConsulta.Add("WHERE BENEFICIARIO LIKE :PARAMETRO")
                vOrdemBusca = 3
              Case 2
                qConsulta.Add("WHERE CODIGOANTIGO LIKE :PARAMETRO")
                vOrdemBusca = 6
              Case 3
                qConsulta.Add("WHERE CODIGODEAFINIDADE LIKE :PARAMETRO")
                vOrdemBusca = 5
            End Select

            If vOrdemBusca <> 0 Then
              qConsulta.Active = False
              qConsulta.ParamByName("PARAMETRO").AsString = TextoBenAtivo + "%"
              qConsulta.Active = True

              If Not qConsulta.FieldByName("HANDLE").IsNull Then
                vContador = 5
              End If
            End If
          End If

          vContador = vContador + 1
        Wend

        If vOrdemBusca = 0 Then
          vOrdemBusca = 3
        End If

        Set qConsulta   = Nothing
        Set qParamBenef = Nothing
      End If

      Set Interface=CreateBennerObject("Procura.Procurar")

      ProcuraBeneficiarioAtivo=Interface.Exec(CurrentSystem,"SAM_BENEFICIARIO|SAM_CONTRATO[SAM_BENEFICIARIO.CONTRATO=SAM_CONTRATO.HANDLE]|SAM_CONVENIO[SAM_BENEFICIARIO.CONVENIO=SAM_CONVENIO.HANDLE]",vColunas,vOrdemBusca,"Matrícula Funcional|Nome|Beneficiario|Contratante|Código Afinidade|Código Antigo|Data Cancelamento|Convenio|Código de origem|Código de repasse",vWhere,"Procura por Beneficiário",False,TextoBenAtivo ,"CA006.ConsultaBeneficiario")

      Set Interface=Nothing
    End If
  End If
End Function


Public Function VERIFICAPAG(psTipoMensagem) As Boolean
  VERIFICAPAG = True
  'calcular a DATA de pagamento se estiver o prestador tambem digitado E ESTIVER EM EDICAO E INSERSAO
  If Not (CurrentQuery.FieldByName("DATARECEBIMENTO").IsNull) Then
    Dim vbVerificarDataPagto As Boolean
    Dim vTabRegimePagtoDLL As Long

    vbVerificarDataPagto = False

    'Em modo Web não existe o componente TABREGIMEPGTO da mesma forma como em modo Desktop
    'Se a macro executar uma referência a este componente em modo Web ocorrerá erro de execução
    If VisibleMode Then
      If (((TABREGIMEPGTO.PageIndex = 0) And _
           (Not CurrentQuery.FieldByName("RECEBEDOR").IsNull)) Or _
          ( _
            (TABREGIMEPGTO.PageIndex = 1) _
            And (   (vTipoPegAnterior <> CurrentQuery.FieldByName("TIPOPEG").AsInteger) _
                 Or (vBeneficiarioAnterior <> CurrentQuery.FieldByName("BENEFICIARIO").AsInteger) _
                 ) _
          ) _
          ) Or (vMudouTabRegimePagto) Or (vMudouDataRecebimentoPEG) Or (vMudouDataPagamentoPEG) Then
        vbVerificarDataPagto = True
        vTipoPegAnterior = CurrentQuery.FieldByName("TIPOPEG").AsInteger
        vBeneficiarioAnterior = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
        VTabRegimePgtoAnterior = CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger
      End If
      vTabRegimePagtoDLL = TABREGIMEPGTO.PageIndex + 1
    Else
      If (((CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) And _
           (Not CurrentQuery.FieldByName("RECEBEDOR").IsNull)) Or _
          (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2)) Then
        vbVerificarDataPagto = True
      End If
      vTabRegimePagtoDLL = CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger
    End If

    If vbVerificarDataPagto Then
      If(CurrentQuery.State =2)Or(CurrentQuery.State =3)Then
        Dim Interface As Object
        Set Interface = CreateBennerObject("SAMCALENDARIOPGTO.ROTINAS")

        'PROCESSO DATA DE PAGAMENTO
        'Caso a dll retorne true,foi encontrado valor,caso false,não foi encontrado data de pagamento.
        datautil = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime

        If Interface.PegarDataPagamento(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("RECEBEDOR").AsInteger, CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime, vTabRegimePagtoDLL, CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger,CurrentQuery.FieldByName("BENEFICIARIO").AsInteger,CurrentQuery.FieldByName("TIPOPEG").AsInteger,CurrentQuery.FieldByName("MUNICIPIO").AsInteger,datautil,AUXNUMPAG)Then

          If (datautil = 0) Or _
             (datautil = -1) Then ' 0 = nao ACHOU /-1 =ACHOU mas estava fechada
            CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0
            ACHOU = False

            If psTipoMensagem <> "" Then 'Se for chamada do UpdateRequired não exibir a mensagem
              bsShowMessage("Data de Pagamento não encontrada ou já processada!", psTipoMensagem)
            End If
            VERIFICAPAG = False
          Else
          	CurrentQuery.FieldByName("DATAPAGAMENTO").Value = datautil
            CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = AUXNUMPAG
          End If
        Else
          CurrentQuery.FieldByName("DATAPAGAMENTO").Clear
          CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0

          ACHOU = False

          If psTipoMensagem <> "" Then 'Se for chamada do UpdateRequired não exibir a mensagem
            bsShowMessage("Data de Pagamento não encontrada ou já processada!", psTipoMensagem)
          End If
          VERIFICAPAG = False
        End If

        Set Interface =Nothing
      End If
    End If
  End If
  vMudouTabRegimePagto = False
  vMudouDataRecebimentoPEG = False
End Function


Public Sub ADIANTAMENTO_OnChange()
  Dim PERCADI As Double
  Dim PERCDESC As Double
  Dim DIAS As Long
  Dim tipoBaseIrrf As String
  Dim CalculaIRRF As String
  Dim mensagem As String

  If processachange2 = False Then
    Exit Sub
  End If

  If TABREGIMEPGTO.PageIndex = 0 Then 'somente para credenciamento
    processachange2 = False
    CurrentQuery.Edit

    If CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "S" Then
      CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "N"
    Else
      CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "S"
    End If

    VerificaRegraPrestador

    CalculaAdiantamento

    processachange2 = True
  End If
End Sub


Public Function TemRegra(PERCADI As Double, PERCDESC As Double, DIAS As Long, mensagem As String, tipoBaseIrrf As String, CalculaIRRF) As Boolean
  Dim PEGDLL As Object
  Set PEGDLL = CreateBennerObject("SAMPEG.ROTINAS")

  TemRegra = False

  If PEGDLL.REGRAADIANTAMENTO(CurrentSystem, CurrentQuery.FieldByName("RECEBEDOR").AsInteger, _
  							  CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime, _
  							  PERCADI, PERCDESC, DIAS, tipoBaseIrrf, CalculaIRRF) Then 'POSSUIR REGRA DE ADIANTAMENTO
    TemRegra = True
  End If

  mensagem = ""

  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger <> 1 Then
    mensagem = "O regime de pagamento deve ser credenciamento" + Chr(13)
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    mensagem = mensagem +"O PEG não pode estar pago" + Chr(13)
  End If

  If Not CurrentQuery.FieldByName("ROTINAFINADIANTAMENTO").IsNull Then
    mensagem = mensagem + "O adiantamento já foi processado"
  End If
End Function

Public Sub PreencherAdiantamentoDigitacaoPeg
  If CurrentQuery.FieldByName("ADIANTAMENTO").AsString <> "S" Then
    processachange2 = True
    ADIANTAMENTO_OnChange
  End If
End Sub

Public Sub CalculaAdiantamento
  processachange = False
  Dim PERCADI As Double
  Dim PERCDESC As Double
  Dim DIAS As Long
  Dim tipoBaseIrrf As String
  Dim CalculaIRRF As String
  Dim mensagem As String

  If CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "S" Then
    If TemRegra(PERCADI, PERCDESC, DIAS, mensagem, tipoBaseIrrf, CalculaIRRF) Then
      If mensagem <> "" Then
        bsShowMessage(mensagem, "I")
      Else
        ComAdiantamento PERCADI, PERCDESC, tipoBaseIrrf, CalculaIRRF
      End If
    Else
      SemAdiantamento 'NAO TEM REGRA DE ADIANTAMENTO
    End If
  Else
    SemAdiantamento
  End If

  processachange = True
End Sub


Public Sub VerificaRegraPrestador
  Dim q1 As Object
  Set q1 =NewQuery

  q1.Clear
  q1.Add("SELECT DATAINICIAL, DATAFINAL FROM SAM_PRESTADOR_ADIANTAMENTO WHERE PRESTADOR = :PRESTADOR")
  q1.ParamByName("PRESTADOR").AsInteger =CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  q1.Active =True

  If q1.FieldByName("DATAINICIAL").IsNull Then
	bsShowMessage("Prestador não possui regra para adiantamento", "A")
	CurrentQuery.FieldByName("ADIANTAMENTO").AsString ="N"
  Else
    If CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime < q1.FieldByName("DATAINICIAL").AsDateTime Or (Not q1.FieldByName("DATAFINAL").IsNull And CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime > q1.FieldByName("DATAFINAL").AsDateTime) Then
	  bsShowMessage("Data de recebimento fora da vigência de regra para adiantamento", "A")
	  CurrentQuery.FieldByName("ADIANTAMENTO").AsString ="N"
    End If
  End If

  Set q1 = Nothing
End Sub


Public Sub VerificaAdiantamento(pMostrarRotulos As Boolean, pCalculaAdiantamento As Boolean)
  If CurrentQuery.FieldByName("ADIANTAMENTO").AsString ="S" Then
    VerificaRegraPrestador

	If pMostrarRotulos Then
		MOSTRAROTULOS
    End If

    If pCalculaAdiantamento Then
		CalculaAdiantamento
    End If

  Else
    bsShowMessage("Necessário marcar a opção adiantamento", "A")
    SemAdiantamento
  End If
End Sub


Public Function DataAdiantamentoOk As Boolean
  If Not CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
    If (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime <= CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime) And _
       (Not CurrentQuery.FieldByName("DATAADIANTAMENTO").IsNull) Then
	  bsShowMessage("A data de Adiantamento deve ser menor que a data de pagamento", "I")
	  DataAdiantamentoOk = False
	  Exit Function
	End If

	If (CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime > CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime) And _
	   (Not CurrentQuery.FieldByName("DATAADIANTAMENTO").IsNull) Then
	  bsShowMessage("A data de Adiantamento deve ser maior ou igual à data de recebimento", "I")
	  DataAdiantamentoOk = False
	  Exit Function
	End If

	If (ServerDate > CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime) And _
	   (Not CurrentQuery.FieldByName("DATAADIANTAMENTO").IsNull) Then
	  bsShowMessage("A data de Adiantamento deve ser maior ou igual à data corrente", "I")
	  DataAdiantamentoOk = False
	  Exit Function
	End If
  End If

  DataAdiantamentoOk = True
End Function


Public Sub ComAdiantamento(PERCADI As Double, PERCDESC As Double, tipoBaseIrrf As String, CalculaIRRF As String)
  If Not DataAdiantamentoOk Then
    Exit Sub
  End If

  Dim diff As Long

  DATAADIANTAMENTO.ReadOnly = AtribuirReadOnly(False)
  VALORADIANTAMENTO.ReadOnly = AtribuirReadOnly(False)
  VALORDESCONTO.ReadOnly = AtribuirReadOnly(False)

  If CurrentQuery.RequestLive Then
	CurrentQuery.Edit

	If gchangeValorAdi Then
	  CurrentQuery.FieldByName("VALORADIANTAMENTO").Value = Arredonda(CurrentQuery.FieldByName("TOTALPAGARINFORMADO").AsFloat * PERCADI /100)
	End If

  If(CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime > CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)Then
	  CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime - 1
	End If

	diff = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime -CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime

	If Arredonda(CurrentQuery.FieldByName("VALORADIANTAMENTO").AsFloat *(PERCDESC /30 /100 * diff)) > 0 Then
	  CurrentQuery.FieldByName("VALORDESCONTO").Value = Arredonda(CurrentQuery.FieldByName("VALORADIANTAMENTO").AsFloat * (PERCDESC /30 /100 * diff))
	Else
	  CurrentQuery.FieldByName("VALORDESCONTO").Value = 0
	End If
  End If

  'calcular a base e o valor do irrf
  Dim vTotalIrrf As Double
  Dim vBaseIRRF As Double

  vTotalIrrf = 0
  vBaseIRRF = 0

  If CalculaIRRF = "S" Then
	Dim vDirf As Long

	If tipoBaseIrrf = "B" Then 'bruto
	  vBaseIRRF = CurrentQuery.FieldByName("VALORADIANTAMENTO").AsFloat
	Else 'liquido
	  vBaseIRRF = CurrentQuery.FieldByName("VALORADIANTAMENTO").AsFloat - CurrentQuery.FieldByName("VALORDESCONTO").AsFloat
	End If
  End If

  CurrentQuery.FieldByName("BASEIRRF").AsFloat = Arredonda(vBaseIRRF)

  MOSTRAROTULOS
End Sub


Public Sub MOSTRAROTULOS
  ROTNUMERODIAS.Text = "Dias de Adiantamento:" + _
  	  Str(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime - CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime)
  TOTALCALCULADO.Text = "Total calculado: " + _
  	  Format(CurrentQuery.FieldByName("VALORADIANTAMENTO").AsFloat - CurrentQuery.FieldByName("VALORDESCONTO").AsFloat,"###,###,##0.00")
End Sub


Public Sub ESCONDEROTULOS
  ROTNUMERODIAS.Text = "Dias de Adiantamento:"
  TOTALCALCULADO.Text = "Total calculado: "
End Sub


Public Sub ARRUMAROTULOS
  If Not CurrentQuery.FieldByName("ROTINAFINADIANTAMENTO").IsNull Then
    DATAADIANTAMENTO.ReadOnly = AtribuirReadOnly(True)
    VALORADIANTAMENTO.ReadOnly = AtribuirReadOnly(True)
    TOTALPAGARINFORMADO.ReadOnly = AtribuirReadOnly(True)
    VALORDESCONTO.ReadOnly = AtribuirReadOnly(True)
    ADIANTAMENTO.ReadOnly = AtribuirReadOnly(True)
  Else
    If CurrentQuery.FieldByName("ADIANTAMENTO").AsBoolean Then
    DATAADIANTAMENTO.ReadOnly = AtribuirReadOnly(False)
    VALORADIANTAMENTO.ReadOnly = AtribuirReadOnly(False)
    TOTALPAGARINFORMADO.ReadOnly = AtribuirReadOnly(False)
    VALORDESCONTO.ReadOnly = AtribuirReadOnly(False)
    ADIANTAMENTO.ReadOnly = AtribuirReadOnly(False)
  End If
  End If

  If CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "S" Then
    MOSTRAROTULOS
  Else
    ESCONDEROTULOS
  End If
End Sub


Public Sub SemAdiantamento
  DATAADIANTAMENTO.ReadOnly = AtribuirReadOnly(True)
  VALORADIANTAMENTO.ReadOnly = AtribuirReadOnly(True)
  VALORDESCONTO.ReadOnly = AtribuirReadOnly(True)

  If CurrentQuery.RequestLive Then
	CurrentQuery.Edit

	CurrentQuery.FieldByName("VALORADIANTAMENTO").Value = 0
	CurrentQuery.FieldByName("VALORDESCONTO").Value = 0
	CurrentQuery.FieldByName("DATAADIANTAMENTO").Clear

	ESCONDEROTULOS
  End If
End Sub


Public Function ContaFinPrestador(Prestador As Long) As Long
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM SFN_CONTAFIN WHERE PRESTADOR=:PRESTADOR AND TABRESPONSAVEL=2")
  sql.ParamByName("PRESTADOR").Value =Prestador
  sql.Active =True

  ContaFinPrestador = sql.FieldByName("HANDLE").AsInteger

  Set sql = Nothing
End Function


Public Function TipoPessoaPrestador(Prestador As Long) As Integer
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE=:PRESTADOR")
  sql.ParamByName("PRESTADOR").Value = Prestador
  sql.Active = True

  TipoPessoaPrestador = sql.FieldByName("FISICAJURIDICA").AsInteger

  Set sql = Nothing
End Function


Public Sub AGENCIA_OnPopup(ShowPopup As Boolean)
  'Não era verificado se a agencia estava ativa ou inativa, exibindo assim todas as agências.
  AGENCIA.LocalWhere = "SITUACAO = 'A'"
End Sub


Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  BENEFICIARIO.LocalWhere = "SAM_BENEFICIARIO.PERMITEREEMBOLSO = 'S'"
End Sub


Public Sub BOTAOALTERARDADOSNF_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  INTERFACE0002.Exec(CurrentSystem, _
                                 1, _
                     "TV_FORM0090", _
            "Dados da Nota Fiscal", _
                                 0, _
                               210, _
                               530, _
                             False, _
                        vsMensagem, _
                       vcContainer)

  Set INTERFACE0002 = Nothing
End Sub

Public Sub BOTAOALTERAREMPENHO_OnClick()
  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    bsShowMessage("É necessário salvar o registro.", "I")
    Exit Sub
  End If

  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  INTERFACE0002.Exec(CurrentSystem, _
					 1, _
		  		     "TV_FORM0150", _
					 "Alterar Empenho", _
					 0, _
		             120, _
					 400, _
					 False, _
		  		     vsMensagem, _
					 vcContainer)

  Set INTERFACE0002 = Nothing
  RefreshPeg
End Sub

Public Sub BOTAOALTERARDATACONTABIL_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  	INTERFACE0002.Exec(CurrentSystem, _
                                   1, _
                       "TV_FORM0083", _
             "Alterar data contabil", _
                                   0, _
                                 170, _
                                 530, _
                               False, _
                          vsMensagem, _
                         vcContainer)

	Set INTERFACE0002 = Nothing
End Sub


Public Sub BOTAOALTERARDATAPAGAMENTO_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  INTERFACE0002.Exec(CurrentSystem, _
                                 1, _
                     "TV_FORM0082", _
          "Alterar data pagamento", _
                                 0, _
                               170, _
                               530, _
                             False, _
                        vsMensagem, _
                       vcContainer)

  Set INTERFACE0002 = Nothing
End Sub


Public Sub BOTAOALTERARDOTACAO_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set vcContainer = NewContainer

  UserVar("HANDLE_PEG") = CurrentQuery.FieldByName("HANDLE").AsString

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  	INTERFACE0002.Exec(CurrentSystem, _
                                   1, _
                       "TV_FORM0147", _
      "Alterar Dotação Orçamentária", _
                                   0, _
                                 320, _
                                 530, _
                               False, _
                          vsMensagem, _
                         vcContainer)

    Set INTERFACE0002 = Nothing
    Set vcContainer = Nothing

    If vsMensagem <> "" Then
      bsShowmessage(vsMensagem, "E")
    Else
      RefreshPeg
    End If
End Sub

Public Sub BOTAOALTERARGUIASAPRESENTADAS_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  	INTERFACE0002.Exec(CurrentSystem, _
                                   1, _
                       "TV_FORM0084", _
           "Alterar quantidade guia", _
                                   0, _
                                 210, _
                                 530, _
                               False, _
                          vsMensagem, _
                         vcContainer)

	Set INTERFACE0002 = Nothing
End Sub


Public Sub BOTAOALTERARIDENTIFICADORPAGTO_OnClick()
  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    bsShowMessage("É necessário salvar o registro.", "I")
    Exit Sub
  End If
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  INTERFACE0002.Exec(CurrentSystem, _
                                 1, _
       "TV_IDENTIFICADORPAGAMENTO", _
      "Alterar identificador de pagamento", _
                                 0, _
                               130, _
                               350, _
                             False, _
                        vsMensagem, _
                       vcContainer)

  Set INTERFACE0002 = Nothing
  RefreshPeg
End Sub

Public Sub BOTAOALTERARVALORAPRESENTADO_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  	INTERFACE0002.Exec(CurrentSystem, _
                                   1, _
                       "TV_FORM0085", _
                     "Alterar valor", _
                                   0, _
                                 210, _
                                 530, _
                               False, _
                          vsMensagem, _
                         vcContainer)

	Set INTERFACE0002 = Nothing
End Sub


Public Sub BOTAOCANCELAFATURAMENTO_OnClick()
  Dim Aux As Boolean

  Dim vFilial As Long

  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial	, "P")

  If Not Aux Then
		AtualizarCarga(True)
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("BSINTERFACE0040.PAGAMENTO")

  Interface.CANCELARPGTOPEG(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set Interface = Nothing

  AtualizarCarga(False)
End Sub


Public Sub BOTAOCANCELARPROVISAO_OnClick()

  If VerificarBloqueioAlteracoesReapresentacao(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG é de reapresentação. ", "I")
	Exit Sub
  End If

  If VerificarBloqueioAlteracoes(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "I")
	Exit Sub
  End If

  Dim retorno As String
  Dim dllCSharp As CSBusinessComponent

  Set dllCSharp = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.Rotinas.Provisionamento.CancelarProvisao,Benner.Saude.Financeiro.Business")
  dllCSharp.AddParameter(pdtInteger,CurrentQuery.FieldByName("HANDLE").AsInteger)
  retorno = dllCSharp.Execute("CancelarPeg")

  If retorno <> "" Then
    bsShowMessage(retorno, "I")
  End If

  Set dllCSharp = Nothing
End Sub


Public Sub BOTAOCOMPENTREGUEOUTROPLANO_OnClick()
  If CurrentQuery.State = 1 Then
    If ((CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2) And _
        (CurrentQuery.FieldByName("POSSUIREEMBOLSOOUTROPLANO").AsString = "S")) Then
      Dim qAlteracao As Object
      Set qAlteracao = NewQuery

      qAlteracao.Clear
      qAlteracao.Add("UPDATE SAM_PEG SET")

      If (CurrentQuery.FieldByName("COMPROVANTEENTREGUEOUTROPLANO").AsString = "N") Then
        qAlteracao.Add("       COMPROVANTEENTREGUEOUTROPLANO = 'S'")
      Else
        qAlteracao.Add("       COMPROVANTEENTREGUEOUTROPLANO = 'N'")
      End If

      qAlteracao.Add(" WHERE HANDLE = :HANDLE")
      qAlteracao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qAlteracao.ExecSQL

      If VisibleMode Then
        SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
      End If

      Set qAlteracao = Nothing
    End If
  End If
End Sub


Public Sub BOTAOGLOSATOTAL_OnClick()
	Dim vsMensagem As String
  	'Verificar se será permitido acionar a funcionalidade se o PEG estiver sendo provisionado
  	If PermissaoAlteracao(0, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, False, vsMensagem) = 1 Then
    	bsShowMessage(vsMensagem, "I")
    	Exit Sub
  	End If

	Dim Interface As Object
	Dim vsMsgErro As String
	Dim viRetorno As Integer

	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
	viRetorno = Interface.Exec(CurrentSystem, _
							   1, _
							   "TV_FORM0098", _
							   "Glosa Total", _
							   0, _
							   290, _
							   516, _
							   False, _
							   vsMsgErro, _
							   Null)

	If viRetorno = 1 Then
	  bsShowMessage(vsMsgErro, "I")
	  Exit Sub
	ElseIf viRetorno = -1 Then
	  Exit Sub
	End If

	Set Interface = Nothing
End Sub


Public Sub BOTAOINCLUIROBS_OnClick()
  If (InStr(SQLServer, "DB2") > 0) And Len(CurrentQuery.FieldByName("OBSERVACAO").AsString) >= 495 Then
    bsShowMessage("O Campo Observação está com a sua capacidade máxima de caracteres, portanto," + Chr(13) + _
		" não será possível incluir observações adicionais.", "E")
    Exit Sub
  End If

  If (InStr(SQLServer, "ORACLE") > 0 Or InStr(SQLServer,"MSSQL") > 0) And Len(CurrentQuery.FieldByName("OBSERVACAO").AsString) >= 4000 Then
    bsShowMessage("O Campo Observação está com a sua capacidade máxima de caracteres, portanto," + Chr(13) + _
		" não será possível incluir observações adicionais.", "I")
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("samPeg.PROCESSAR")
  Interface.IncluirObs(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface =Nothing
  AtualizarCarga(False)
End Sub


Public Sub BOTAOINCLUIRPRESTADOR_OnClick()
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


Public Sub BOTAOLIBERARVERIFICACAO_OnClick()
  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor, confirme ou cancele as alterações!", "I")
    Exit Sub
  End If

  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean

  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  Dim Interface As Object
  If WebMode Then
  	Dim retorno As Long
	Dim vsMensagemErro As String
	Dim vcContainer As CSDContainer
	Set vcContainer = NewContainer
	vcContainer.AddFields("HANDLE:INTEGER")

	vcContainer.Insert
	vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
	retorno = Interface.ExecucaoImediata(CurrentSystem, _
			  "BSPro000", _
			  "LiberaVerificacao", _
			  "Liberar Verificação", _
			  CurrentQuery.FieldByName("HANDLE").AsInteger, _
			  "SAM_PEG", _
			  "SITUACAOPROCESSAMENTO", _
			  "", _
			  "", _
			  "P", _
			  True, _
			  vsMensagemErro, _
			  vcContainer)
	If retorno = 0 Then
		bsShowMessage("Processo enviado para execução no servidor!", "I")
		bsShowMessage("Verificando informações para Liberação!", "I")

	Else
		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	End If
  Else
	Set Interface = CreateBennerObject("BSPro000.Rotinas")

    	Interface.LiberaVerificacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	WriteAudit("M",HandleOfTable("SAM_PEG"),CurrentQuery.FieldByName("HANDLE").AsInteger,"PEG - Libera verificação")
	AtualizarCarga(False)
  End If
  Set Interface = Nothing
End Sub

Public Sub BOTAOPEGORIGINAL_OnClick()
  If CurrentQuery.FieldByName("PEGORIGINAL").AsInteger > 0 Then
    Dim Interface As Object
    Set Interface = CreateBennerObject("BSINTERFACE0050.ROTINAS")

    Interface.MontarForm(CurrentSystem, "SAM_PEG", CurrentQuery.FieldByName("PEGORIGINAL").AsInteger, "Peg Original")
    Set Interface = Nothing
  End If
End Sub


Public Sub BOTAOPROVISIONARPEG_OnClick()

  If VerificarBloqueioAlteracoesReapresentacao(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG é de reapresentação. ", "I")
	Exit Sub
  End If

  If VerificarBloqueioAlteracoes(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "I")
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("PRONTOPARAPROVISAO").AsString <> "S") Then
    bsShowMessage("O PEG não está pronto para provisão!", "I")
    Exit Sub
  End If

  Dim retorno As String
  Dim dllCSharp As CSBusinessComponent

  Set dllCSharp = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.Rotinas.Provisionamento.ProvisionarPeg,Benner.Saude.Financeiro.Business")
  dllCSharp.AddParameter(pdtInteger,CLng(CurrentQuery.FieldByName("HANDLE").AsInteger))
  retorno = dllCSharp.Execute("Provisionar")

  If retorno <> "" Then
    bsShowMessage(retorno, "I")
  End If

  Set dllCSharp = Nothing
End Sub


Public Sub BOTAORECLASSIFICAR_OnClick()
  Dim Aux As Boolean
  Dim vsMensagem As String

  'verifica se há fatura de provisão contabilizada
  Dim CheckFatProvisao As CSBusinessComponent

  Set CheckFatProvisao = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.Rotinas.Provisionamento.CancelarProvisao, Benner.Saude.Financeiro.Business")

  CheckFatProvisao.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  vsMensagem = CheckFatProvisao.Execute("VerificaFaturaProvisaoPorPeg")

  If(vsMensagem <> "") Then
  	bsShowMessage(vsMensagem, "E")
	Set CheckFatProvisao = Nothing
  	Exit Sub
  End If

  Set CheckFatProvisao = Nothing

 'Verificar se é permitido excluir o PEG conforme as regras do Provisionamento
 If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem) = 1 Then
    bsShowMessage(vsMensagem, "E")
    Exit Sub
 End If

  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If VisibleMode = True Then
    Dim Interface As Object

    Set Interface = CreateBennerObject("BSINTERFACE0058.Rotinas")

    Interface.Reclassificar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    Dim vsMensagemErro As String
    Dim Obj As Object
    Dim viRet As Long
    Dim vcContainer As CSDContainer
   	Set vcContainer = NewContainer
   	vcContainer.AddFields("HANDLE:INTEGER")

	vcContainer.Insert
 	vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRet = Obj.ExecucaoImediata(CurrentSystem, _
		                                	 "SamPeg", _
		                                	 "ReclassificarPEG", _
		                                	 "Reclassificar Guias", _
		                                	 CurrentQuery.FieldByName("HANDLE").AsInteger, _
		                                	 "SAM_PEG", _
		                                	 "SITUACAORECLASSIFICACAO", _
		                                	 "", _
		                                	 "", _
		                                	 "P", _
		                                	 True, _
		                                	 vsMensagemErro, _
		                                	 vcContainer)

			If viRet = 0 Then
			 	bsShowMessage("Processo enviado para execução no servidor!", "I")
			Else
		     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		   	End If

			Set Obj = Nothing

  End If
  Set Interface = Nothing
End Sub


Public Sub BOTAOVERIFICAMONITORAMENTO_OnClick()
  Dim INTERFACE0002 As Object
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  IncluiSessionVarMonitoramento

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
    INTERFACE0002.Exec(CurrentSystem, _
                       1, _
                       "TV_FORM0091", _
                       "Dados Inconsistentes para o Monitoramento", _
                       0, _
                       400, _
                       420, _
                       False, _
                       vsMensagem, _
                       vcContainer)

  Set INTERFACE0002 = Nothing
End Sub


Public Sub CONCILIARNOTA_OnClick()
  Dim q1 As Object
  Set q1 = NewQuery

  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If Not(CurrentQuery.FieldByName("RECEBEDOR").IsNull) And _
     Not(CurrentQuery.FieldByName("NFNUMERO").IsNull) Then
    q1.Clear
    q1.Add("SELECT DISTINCT(CF.HANDLE) CONTAFINANCEIRA")
    q1.Add("  FROM SFN_CONTAFIN  CF,")
    q1.Add("       SFN_NOTA      NOTA")
    q1.Add(" WHERE CF.HANDLE   = NOTA.CONTAFINANCEIRA")
    q1.Add("   AND CF.PRESTADOR = :PRESTADOR")

    q1.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    q1.Active = True

    If (q1.FieldByName("CONTAFINANCEIRA").IsNull) Then
      bsShowMessage("Não foi encontrado a determinada nota para este prestador", "I")
      Exit Sub
    Else
      Dim q2 As Object
      Set q2 = NewQuery

      q2.Clear
      q2.Add("SELECT HANDLE                   ")
      q2.Add("  FROM SFN_NOTA                 ")
      q2.Add(" WHERE CONTAFINANCEIRA = :CONTA ")
      q2.Add("   AND NUMERO = :NOTA           ")

      q2.ParamByName("CONTA").AsInteger = q1.FieldByName("CONTAFINANCEIRA").AsInteger
      q2.ParamByName("NOTA").AsString = CurrentQuery.FieldByName("NFNUMERO").AsString
      q2.Active = True

      If Not(q2.FieldByName("HANDLE").IsNull) Then
        q1.Clear
        q1.Add("UPDATE SAM_PEG SET NF = :NF WHERE HANDLE = :HANDLE")

        q1.ParamByName("NF").AsInteger = q2.FieldByName("HANDLE").AsInteger
        q1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

        q1.ExecSQL

        bsShowMessage("Nota encontrada e conciliada", "I")
      Else
        bsShowMessage("Não foi encontrado a determinada nota para este prestador", "I")
        Exit Sub
      End If
    End If
  End If

  Set q1 = Nothing
  Set q2 = Nothing
End Sub


Public Sub CONTATERCEIRO_OnClick()
  If (CurrentQuery.State <> 1) Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If (CurrentQuery.FieldByName("SITUACAO").AsString = "1") Then
	Dim Interface As Object
	Set Interface = CreateBennerObject("BsPro006.Rotinas")

	Interface.ContaTerceiro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	Set Interface = Nothing

    If VisibleMode Then
	  SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
	End If
  Else
	bsShowMessage("Os dados da conta de terceiro só podem ser alterados em fase de digitação.", "I")
  End If
End Sub


Public Sub CRITICARDIGITACAO_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "2" Or CurrentQuery.FieldByName("SITUACAO").AsString = "3" Or _
	CurrentQuery.FieldByName("SITUACAO").AsString = "4"         Then
    bsShowMessage("Comando válido somente na fase de digitação", "I")
    Exit Sub
  End If

  If CurrentQuery.State<>1 Then
	bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim ConsultaSqlGuia As Object
  Dim ConsultaSqlEvento As Object
  Dim ListaSqlEvento As Object
  Dim TotalGuiaDigitada As Integer
  Dim TotalGuiaPEG As Integer
  Dim TotalGuiaEvento As Integer
  Dim Diferenca As Integer
  Dim Mensagem As String
  Dim RelGuias As String
  Dim PrimeiraGuia As Long
  Set ConsultaSqlGuia = NewQuery
  Set ConsultaSqlEvento = NewQuery
  Set ListaSqlEvento = NewQuery

  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem,vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  ConsultaSqlGuia.Add("SELECT COUNT(G.GUIA)    QUANT        ")
  ConsultaSqlGuia.Add("FROM  SAM_PEG     P")
  ConsultaSqlGuia.Add("JOIN  SAM_GUIA   G ON P.HANDLE = G.PEG ")
  ConsultaSqlGuia.Add("WHERE P.HANDLE = :PEG")

  ConsultaSqlGuia.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  ConsultaSqlGuia.Active = True

  TotalGuiaPEG = CurrentQuery.FieldByName("QTDGUIA").AsInteger
  TotalGuiaDigitada = ConsultaSqlGuia.FieldByName("QUANT").AsInteger

  ' Mochi Cabesp INICIO
  Dim qContaGuiasDevolvidas As Object
  Set qContaGuiasDevolvidas = NewQuery

  qContaGuiasDevolvidas.Clear
  qContaGuiasDevolvidas.Add("SELECT COUNT(*) DEVOLVIDA FROM SAM_GUIA_DEVOLUCAO WHERE PEG = :PEG")

  qContaGuiasDevolvidas.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qContaGuiasDevolvidas.Active = True

  TotalGuiaDigitada = TotalGuiaDigitada + qContaGuiasDevolvidas.FieldByName("DEVOLVIDA").AsInteger

  Set qContaGuiasDevolvidas = Nothing

  If TotalGuiaDigitada > CurrentQuery.FieldByName("QTDGUIA").AsInteger Then
  ' Mochi Cabesp FIM
    bsShowMessage("Quantidade de guias digitadas " + Str(TotalGuiaDigitada) + " diverge do total informado no PEG " + _
    	Str(TotalGuiaPEG), "I")

    Set ConsultaSqlGuia = Nothing

    Exit Sub
  End If

  ' Mochi Cabesp INICIO
  Diferenca = (CurrentQuery.FieldByName("QTDGUIA").AsInteger - TotalGuiaDigitada)
  ' Mochi Cabesp FIM

  Mensagem = ""

  ' Mochi Cabesp INICIO
  If TotalGuiaDigitada < CurrentQuery.FieldByName("QTDGUIA").AsInteger Then
  ' Mochi Cabesp FIM
    Mensagem = "Quantidade de guias que faltam ser digitadas: " + Str(Diferenca)

    Set ConsultaSqlGuia = Nothing
  End If

  ConsultaSqlEvento.Add("SELECT COUNT(G.GUIA)    QUANT")
  ConsultaSqlEvento.Add("     FROM  SAM_PEG          P")
  ConsultaSqlEvento.Add("     JOIN  SAM_GUIA         G  ON P.HANDLE =G.PEG             ")
  ConsultaSqlEvento.Add("WHERE P.HANDLE = :PEG")
  ConsultaSqlEvento.Add("AND G.PEG=P.HANDLE AND NOT EXISTS (SELECT GE.HANDLE FROM SAM_GUIA_EVENTOS GE WHERE GE.GUIA=G.HANDLE) ")

  ConsultaSqlEvento.ParamByName("PEG").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  ConsultaSqlEvento.Active= True

  TotalGuiaEvento = ConsultaSqlEvento.FieldByName("QUANT").AsInteger

  If ConsultaSqlEvento.FieldByName("QUANT").AsInteger > 0 Then
    Mensagem = Mensagem + Chr(13)+Chr(10) + "Quantidade de guias SEM EVENTOS: " + Str(TotalGuiaEvento) + Chr(13) + Chr(10) + _
    	"Ordem" + "   Guia" + Chr(13) + Chr(10)
    Set ConsultaSqlEvento = Nothing
  End If

  Set ConsultaSqlGuia = Nothing
  Set ConsultaSqlEvento = Nothing

  If Mensagem <> "" Then
    If WebMode Then
      bsShowMessage(Mensagem, "I")
    Else
      Begin Dialog UserDialog 470,231 ' %GRID:10,7,1,1
        OKButton 355,30,90,21
        Text 15,7,340,14,"Problemas identificados no PEG",.Text1
        TextBox 10,60,450,112,.TextBox1,1
      End Dialog

      Dim dlg As UserDialog

      dlg.TextBox1 = ""

      ListaSqlEvento.Add("SELECT G.GUIA    GUIA, G.ORDEM ORDEM, G.HANDLE GHANDLE")
      ListaSqlEvento.Add("     FROM  SAM_PEG          P")
      ListaSqlEvento.Add("     JOIN  SAM_GUIA        G  ON P.HANDLE = G.PEG               ")
      ListaSqlEvento.Add("WHERE P.HANDLE = :PEG")
      ListaSqlEvento.Add("AND G.PEG=P.HANDLE AND NOT EXISTS (SELECT GE.HANDLE FROM SAM_GUIA_EVENTOS GE WHERE GE.GUIA=G.HANDLE) ")
      ListaSqlEvento.Add("ORDER BY G.ORDEM")

      ListaSqlEvento.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      ListaSqlEvento.Active= True

      RelGuias = ""

      ListaSqlEvento.First

      PrimeiraGuia = ListaSqlEvento.FieldByName("GHANDLE").AsInteger

      While(Not(ListaSqlEvento.EOF))
        RelGuias = RelGuias + "  " + Str(ListaSqlEvento.FieldByName("ORDEM").AsInteger) + "       " + _
      	  Str(ListaSqlEvento.FieldByName("GUIA").AsFloat) + Chr(13) + Chr(10)

        ListaSqlEvento.Next
      Wend

      Set ListaSqlEvento = Nothing

      dlg.TextBox1 = Mensagem + RelGuias

      Dialog dlg


      Dim Interface3 As Object
      Set Interface3 = CreateBennerObject("BSPRO006.ROTINAS")

      Interface3.Digitar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)

      Set Interface3 =Nothing

    End If
  Else
    bsShowMessage("PEG sem problemas para Mudar de Fase", "I")
  End If
End Sub


Public Sub DATAPAGAMENTO_OnChange()
  If Not(CurrentQuery.FieldByName("NUMEROPAGAMENTO").AsInteger > 0)Then
    CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0
  End If
End Sub


Public Sub DATARECEBIMENTO_OnChange()
  If(CurrentQuery.State <> 3)Then
  	vMudouDataRecebimentoPEG = True
  End If
End Sub


Public Sub DESDOBRAR_OnClick()

  If VerificarBloqueioAlteracoesReapresentacao(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG é de reapresentação. ", "I")
	Exit Sub
  End If

  If VerificarBloqueioAlteracoes(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("Esta ação não pode ser realizada porque o PEG está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "I")
	Exit Sub
  End If

  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If CurrentQuery.FieldByName("PEGORIGINAL").AsInteger > 0 Then
    bsShowMessage("Não é possível desdobrar um PEG de reapresentação", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString ="4" Then
    bsShowMessage("PEG já foi pago", "I")
    Exit Sub
  End If

  If VisibleMode Then
  	Dim Interface As Object
  	Dim vsMsg As String
  	Dim viRetorno As Long

  	If CurrentQuery.FieldByName("SITUACAO").AsString <> "3" Then
      'Set Interface = CreateBennerObject("SAMPEGDESDOBRA.PROCESSOS")
  	  Set Interface = CreateBennerObject("BSINTERFACE0046.ROTINAS")
      'ViRetorno = Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMsg)
	  viRetorno = Interface.DesdobrarPEG(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsMsg)
	End If

  	Set Interface = Nothing

	If viRetorno = 0 Then
  		If vsMsg <> "" Then
  			bsShowMessage(vsMsg, "E")
  			Exit Sub
  		End If
  	Else
    	bsShowMessage(vsMsg + Chr(13) + "Desdobramento concluido", "I")
    End If

  Else

	Dim vsMensagemErro As String
  	Dim viRet As Long
  	Dim Obj As Object
    Dim vcContainer As CSDContainer

    Set vcContainer = NewContainer
        vcContainer.AddFields("HANDLE:INTEGER")
        vcContainer.Insert
        vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			viRet = Obj.ExecucaoImediata(CurrentSystem, _
		                                	 "SAMPEGDESDOBRA", _
		                                	 "Processos", _
		                                	 "Desdobramento de peg", _
		                                	 0, _
		                                	 "", _
		                                	 "", _
		                                	 "", _
		                                	 "", _
		                                	 "P", _
		                                	 False, _
		                                	 vsMensagemErro, _
		                                	 vcContainer)

			If viRet = 0 Then
			 	bsShowMessage("Processo enviado para execução no servidor!", "I")
			Else
		     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		   	End If

			Set Obj = Nothing
  End If

  AtualizarCarga(False)
End Sub


Public Sub DETALHESBENEFICIARIO_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("CA006.ConsultaBeneficiario")
  Interface.info(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, , ,1)

End Sub


Public Sub DEVOLVERPEG_OnClick()
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim vsMensagem As String

  'Verificar se é permitido excluir o PEG conforme as regras do Provisionamento
  If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem) = 1 Then
    bsShowMessage(vsMensagem, "E")
    Exit Sub
  End If

  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If
  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
	AtualizarCarga(True)
    Exit Sub
  End If
  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("Peg já devolvido", "I")
    Exit Sub
  End If

  If (Not CurrentQuery.FieldByName("AGRUPADORPAGAMENTO").IsNull) Then
  	BsShowMessage("Não é permitida a devolução de PEG ligado a registro de pagamento.","I")
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("SamDevolucaoGuia.Rotinas")

  Interface.DevolverPeg("", "S", CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set Interface =Nothing
  AtualizarCarga(False)

End Sub


Public Sub BOTAODIGITAR_OnClick()
  Dim Aux As Boolean

  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
    If CurrentQuery.State <> 1 Then
      bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) And (CurrentQuery.FieldByName("RECEBEDOR").IsNull) Then
      bsShowMessage("Recebedor é obrigatório para PEG de credenciamento", "I")
      Exit Sub
    End If

    If PermitePeloGrupoSeguranca("SAM_PEG", "BOTAODIGITAR") Then

      ' TRATAMENTO PARA NÃO DEIXAR DIGITAR GUIAS ENQUANTO NÃO FOR RODADO O AGENDAMENTO TISS (CASO PEG SEJA ORIUNDO DE IMPORTAÇÃO TISS
      If Not PegVinculadoImportacaoTISS(" AND SITUACAO = 'A'" ) Then
      	Dim Interface As Object
      	Set Interface = CreateBennerObject("BSPRO006.ROTINAS")
	    Interface.Digitar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)
    	Set Interface = Nothing
      Else
        bsShowMessage("Comando não permitido em peg criado via importação tiss, cujo processamento do agendamento não tenha sido realizado", "I")
      End If

    Else
      bsShowMessage("Comando não permitido pelo grupo de segurança", "I")
    End If

  Else
    bsShowMessage("Comando válido somente na fase de digitação", "I")
  End If

  If VisibleMode Then
    SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If
End Sub


Public Sub CONFERIDO_OnClick()
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString ="1" Then 'digitacao
    bsShowMessage("PEG ainda em digitacao", "I")
  ElseIf CurrentQuery.FieldByName("SITUACAO").AsString ="4" Then 'pago
    bsShowMessage("PEG já foi pago", "I")
  Else

    Dim sql2 As Object
    Set sql2 = NewQuery

    sql2.Add("SELECT DISTINCT GE.GUIA     ")
    sql2.Add("  FROM SAM_GUIA_EVENTOS GE, ")
    sql2.Add("       SAM_GUIA G,          ")
    sql2.Add("       SAM_PEG PG           ")
    sql2.Add(" WHERE GE.GUIA = G.HANDLE   ")
    sql2.Add("   AND G.PEG = PG.HANDLE    ")
    sql2.Add("   AND GE.SITUACAO <= '2'   ")
    sql2.Add("   AND PG.HANDLE = " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger))
    sql2.Add("   AND GE.COPIAEVENTOORIGINAL <> 'S' ")
    sql2.Active = True

    If Not sql2.FieldByName("GUIA").IsNull Then
      bsShowMessage("PEG com guias Pendentes", "I")
      Exit Sub
    End If

    If Not InTransaction Then StartTransaction

    Dim sql As Object
    Set sql =NewQuery

    sql.Add("UPDATE SAM_GUIA SET USUARIOCONFERENTE=:USU, DATACONFERENCIA=:DATA WHERE PEG=:PEG")

    sql.ParamByName("USU").Value = CurrentUser
    sql.ParamByName("DATA").Value = ServerDate
    sql.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

    sql.ExecSQL
    sql.Clear

    sql.Add("UPDATE SAM_GUIA_EVENTOS SET USUARIOCONFERENTE=:USU,DATACONFERENCIA=:DATA ")
    sql.Add("WHERE GUIA IN (SELECT HANDLE FROM SAM_GUIA WHERE PEG=:PEG) AND COPIAEVENTOORIGINAL <> 'S' ") 'Coelho SMS: 96895

    sql.ParamByName("USU").Value = CurrentUser
    sql.ParamByName("DATA").Value = ServerDate
    sql.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

    sql.ExecSQL

    If InTransaction Then Commit

    If VisibleMode Then
      SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
    End If
  End If
End Sub


Public Sub EXCLUIRPEG_OnClick()
  If CurrentQuery.State <> 1 Then
  	bsShowMessage("O PEG está em edição! Por favor, confirme ou cancele as alterações!", "I")
    Exit Sub
  End If

  Dim vsMensagem As String

  'Verificar se é permitido excluir o PEG conforme as regras do Provisionamento
  If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem) = 1 Then
    bsShowMessage(vsMensagem, "I")
    Exit Sub
  End If

  If PegVinculadoImportacaoTISS("") Then
    bsShowMessage("Peg vinculado a uma importação TISS, somente é possível a sua devolução", "I")
    Exit Sub
  End If

  Dim vContinue As Boolean
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT OG.CODIGO FROM SAM_PEG P JOIN SAM_GUIA G ON (G.PEG = P.HANDLE) ")
  sql.Add("								   JOIN SIS_ORIGEMGUIA OG ON (G.ORIGEMGUIA = OG.HANDLE)")
  sql.Add("               WHERE P.HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  sql.Active = True

  If (sql.FieldByName("CODIGO").AsInteger = 3) Or (sql.FieldByName("CODIGO").AsInteger = 4) Then
  	bsShowMessage("Não é possivel excluir PEGs gerados pelo autorizador externo", "I")
  	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("Não é permita a exclusão de Peg's faturados", "I")
    Exit Sub
  End If


  If (Not CurrentQuery.FieldByName("AGRUPADORPAGAMENTO").IsNull) Then
  	BsShowMessage("Não é permitida a exclusão de PEG ligado a registro de pagamento.","I")
    Exit Sub
  End If

  Dim qBusca As Object
  Dim vsPrimeiraPassagem As String

  Set qBusca = NewQuery

  qBusca.Clear
  qBusca.Add("SELECT COUNT(HANDLE) CONT FROM SAM_GUIA WHERE PEG = :HANDLEPEG")
  qBusca.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qBusca.Active = True

  If bsShowMessage("Deseja excluir este PEG juntamente com a(s) sua(s) guia(s) : " + qBusca.FieldByName("CONT").AsString,"Q") = vbYes Then
    'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
    Dim vFilial As Long
    If (VisibleMode Or WebMode) Then
	  vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
    Else
	  vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
    End If
    vContinue = CheckFilialProcessamento(CurrentSystem, vFilial, "E")

    If vContinue = False Then
      AtualizarCarga(True)
      Exit Sub
    End If

    qBusca.Clear
    qBusca.Add("SELECT COUNT(HANDLE) CONT FROM SAM_GUIA_DEVOLUCAO WHERE PEG = :HANDLEPEG")
    qBusca.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qBusca.Active = True

    If (qBusca.FieldByName("CONT").AsInteger > 0) Then
      bsShowMessage("Impossível remover este PEG, pois este possui GUIA(s) devolvida(s).", "I")
      Set qBusca = Nothing
      Exit Sub
    Else
      qBusca.Clear
      qBusca.Add("SELECT HANDLE FROM SFN_ROTINAFINPAG WHERE PEGINICIAL = :HANDLEPEG OR PEGFINAL = :HANDLEPEG")
      qBusca.Add("UNION ")
      qBusca.Add("SELECT HANDLE FROM SFN_ROTINAFINPAG_PEG WHERE PEG = :HANDLEPEG")
      qBusca.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qBusca.Active = True

      If Not(qBusca.EOF) Then
        bsShowMessage("Impossível remover este PEG, pois o mesmo está relacionado à uma rotina de pagamento.", "I")
        Set qBusca = Nothing
        Exit Sub
      Else
        Dim Interface As Object

        If Not WebMode Then
          Set Interface =CreateBennerObject("SAMPEG.processar")
          Interface.DeletarPeg(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
          bsShowMessage("Peg excluído com sucesso!", "I")
        Else

          Dim vsMensagemErro As String
          Dim Obj As Object
          Dim viRet As Long
          Dim vcContainer As CSDContainer
          Set vcContainer = NewContainer
       	  vcContainer.AddFields("HANDLE:INTEGER")

          vcContainer.Insert
       	  vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
		  viRet = Obj.ExecucaoImediata(CurrentSystem, _
	                               	  "SamPeg", _
	                               	  "Deletar_Peg", _
	                               	  "Exclusão de Peg", _
	                               	  CurrentQuery.FieldByName("HANDLE").AsInteger, _
		                              "SAM_PEG", _
		                              "SITUACAOPROCESSAMENTO", _
		                              "", _
		                              "", _
		                              "P", _
		                              True, _
		                              vsMensagemErro, _
		                              vcContainer)

			If viRet = 0 Then
			 	bsShowMessage("Processo enviado para execução no servidor!", "I")
			 	bsShowMessage("Peg sendo excluído!", "I")

			Else
		     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
		   	End If

			Set Obj = Nothing

        End If
        Set Interface = Nothing
      End If
    End If

  End If

  Set qBusca = Nothing
  AtualizarCarga(False)
End Sub


Public Sub FASEPEGTODOS_OnClick()
  Dim Interface As Object
  Dim Aux As Boolean
  Dim vFilial As Long

  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  Set Interface = CreateBennerObject("BSINTERFACE0039.ProcessarPegLote")

  Interface.Exec(CurrentSystem)

  Set Interface =Nothing

  AtualizarCarga(False)
End Sub


Public Sub FASEPEG_OnClick()
  Dim qFaseAtual As String

  If (CurrentQuery.FieldByName("SITUACAO").AsString = "2") Then
    Set qVerificaConsiderarSp = NewQuery
    Set qServicosPrestador = NewQuery
    Set qPreencheServicos = NewQuery

    qVerificaConsiderarSp.Active = False
    qVerificaConsiderarSp.Clear
    qVerificaConsiderarSp.Add("SELECT CONSIDERARCODSERVICO FROM SFN_PARAMETROSFIN")
    qVerificaConsiderarSp.Active = True

    qServicosPrestador.Active = False
    qServicosPrestador.Clear
    qServicosPrestador.Add("SELECT CODIGOSERVICO, CODIGOSERVICOPREFSP FROM SAM_PRESTADOR WHERE HANDLE = :PPRESTHANDLE")
    qServicosPrestador.ParamByName("PPRESTHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    qServicosPrestador.Active = True

    If (qVerificaConsiderarSp.FieldByName("CONSIDERARCODSERVICO").AsString = "S") And (CurrentQuery.FieldByName("TABREGIMEPGTO").Value = "1") Then
      If ((CurrentQuery.FieldByName("LISTASERVICO").IsNull) And (CurrentQuery.FieldByName("CODIGOSERVICO").IsNull)) Then
        If ((Not qServicosPrestador.FieldByName("CODIGOSERVICO").IsNull) And (Not qServicosPrestador.FieldByName("CODIGOSERVICOPREFSP").IsNull)) Then
          qPreencheServicos.Clear
          qPreencheServicos.Add("UPDATE SAM_PEG SET LISTASERVICO = :PLISTASERVICO, CODIGOSERVICO = :PCODIGOSERVICO WHERE HANDLE = :PHANDLE")
          qPreencheServicos.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          qPreencheServicos.ParamByName("PLISTASERVICO").AsInteger = qServicosPrestador.FieldByName("CODIGOSERVICO").AsInteger
          qPreencheServicos.ParamByName("PCODIGOSERVICO").AsInteger = qServicosPrestador.FieldByName("CODIGOSERVICOPREFSP").AsInteger
          qPreencheServicos.ExecSQL
        Else
          bsshowmessage("Necessário informar um serviço no Prestador Recebedor ou no PEG","E")
          Exit Sub
        End If
      ElseIf ((CurrentQuery.FieldByName("LISTASERVICO").IsNull) Or (CurrentQuery.FieldByName("CODIGOSERVICO").IsNull)) Then
        bsshowmessage("Necessário informar um serviço no Prestador Recebedor ou no PEG","E")
        Exit Sub
      End If
    End If
  End If

  qFaseAtual = CurrentQuery.FieldByName("SITUACAO").AsString

  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode mudar de fase", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode mudar de fase", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode mudar de fase", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  'Luciano T. Alberti - SMS 64356 - 05/09/2007 - Início
  If (qFaseAtual = "2") And (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) Then

    Dim dll As Object
    Dim vbMudaFAsePegComNF As Boolean

    Set dll = CreateBennerObject("Especifico.uEspecifico")

	vbMudaFAsePegComNF = dll.PRO_VerificaDadosNotaFiscalDoPeg(CurrentSystem, CurrentQuery.FieldByName("RECEBEDOR").AsInteger, CurrentQuery.FieldByName("NFNUMERO").AsString)

	If Not(vbMudaFAsePegComNF) Then
	  Set dll =Nothing
	  Exit Sub
	End If

	Set dll =Nothing
  End If

 'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If
  Aux = CheckFilialProcessamento(CurrentSystem,vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If


  Dim qAjustaDataModificacao As Object
  Set qAjustaDataModificacao = NewQuery

  qAjustaDataModificacao.Clear
  qAjustaDataModificacao.Add("UPDATE SAM_PEG ")
  qAjustaDataModificacao.Add("   SET DATA = :DATA")
  qAjustaDataModificacao.Add(" WHERE HANDLE = :HANDLE")
  qAjustaDataModificacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qAjustaDataModificacao.ParamByName("DATA").AsDateTime = ServerNow

  If Not InTransaction Then StartTransaction
  qAjustaDataModificacao.ExecSQL
  If InTransaction Then Commit

  Set qAjustaDataModificacao = Nothing

  VERIFICAPAG("I")

  If CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
    bsShowMessage("Campo Pagamento inválido", "I")
    Exit Sub
  End If

  Dim FASE As Long
  Dim OldFase As Long
  Dim Interface As Object

  FASE = CurrentQuery.FieldByName("SITUACAO").AsInteger
  OldFase = FASE

  If VisibleMode Then
  	Set Interface = CreateBennerObject("BSINTERFACE0046.Rotinas")
	Interface.MudarFase(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  	WriteAudit("M", HandleOfTable("SAM_PEG"), CurrentQuery.FieldByName("HANDLE").AsInteger, "PEG - Mudança de Fase - fase:" + Str(FASE))
  Else
    Dim qPegMudandoDeFase As Object
    Set qPegMudandoDeFase = NewQuery
    qPegMudandoDeFase.Clear
    qPegMudandoDeFase.Add("UPDATE SAM_PEG SET PEGMUDANDODEFASE = 'S' WHERE HANDLE = :HANDLE")
    qPegMudandoDeFase.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qPegMudandoDeFase.ExecSQL


	Dim vsMensagemErro As String
  	Dim viRet As Long
  	Dim Obj As Object

	Dim vcContainer As CSDContainer
   	Set vcContainer = NewContainer
   	vcContainer.AddFields("HANDLE:INTEGER")

	vcContainer.Insert
 	vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRet = Obj.ExecucaoImediata(CurrentSystem, _
                                	 "BSPro000", _
                                	 "Rotinas", _
                                	 "Mudança de fase do peg", _
                                	 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                	 "SAM_PEG", _
                                	 "SITUACAOPROCESSAMENTO", _
                                	 "", _
                                	 "", _
                                	 "P", _
                                	 True, _
                                	 vsMensagemErro, _
                                	 vcContainer)

	If viRet = 0 Then
	 	bsShowMessage("Processo enviado para execução no servidor!", "I")
	Else
     	bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
   	End If

	Set Obj = Nothing
  End If

  If VisibleMode Then
	SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If

  If FASE <> OldFase Then
    InicializaSamParametrosProcContas("UTILIZAHISTORICOODONTOLOGICO")
    If qSamParametrosProcContas.FieldByName("UTILIZAHISTORICOODONTOLOGICO").AsString = "S" Then
      Dim SQLGUIAS As Object
      Set SQLGUIAS = NewQuery
      Dim SQLEVENTOS As Object
      Set SQLEVENTOS = NewQuery

      If (CurrentQuery.FieldByName("SITUACAO").AsString = "2") And _
         (OldFase = 1) Then 'SE FOR DE DIGITAÇÃO PARA CONFERÊNCIA,VERIFICA HISTÓRICO
        Dim RESULTADO As Boolean
        Dim AUDITORIA As Object
        Set AUDITORIA = CreateBennerObject("BSCLI006.ROTINAS")

        SQLGUIAS.Add("SELECT GUIA.HANDLE, ")
        SQLGUIAS.Add("       GUIA.BENEFICIARIO ")
        SQLGUIAS.Add("  FROM SAM_GUIA GUIA,")
        SQLGUIAS.Add("       SAM_TIPOGUIA_MDGUIA MDGUIA,")
        SQLGUIAS.Add("       SAM_TIPOGUIA TIPOGUIA,")
        SQLGUIAS.Add("       SAM_PEG PEG")
        SQLGUIAS.Add(" WHERE GUIA.MODELOGUIA = MDGUIA.HANDLE")
        SQLGUIAS.Add("   AND MDGUIA.TIPOGUIA = TIPOGUIA.HANDLE")
        SQLGUIAS.Add("   AND TIPOGUIA.TABTIPOGUIA = 3")
        SQLGUIAS.Add("   AND GUIA.PEG = PEG.HANDLE")
        SQLGUIAS.Add("   AND PEG.HANDLE = :PEG")

        SQLGUIAS.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        SQLGUIAS.Active = True

        While Not SQLGUIAS.EOF
          RESULTADO = AUDITORIA.AUDITAR(CurrentSystem, SQLGUIAS.FieldByName("HANDLE").AsInteger)

          SQLGUIAS.Next
        Wend
      End If
    End If
	FinalizaSamParametrosProcContas

    Set SQLGUIAS =Nothing

    'Mudando de digitação para pronto
    If qFaseAtual = "1" Then
      Dim qProblema As Object
      Dim qAltera As Object
      Set qProblema = NewQuery
      Set qAltera = NewQuery

      qProblema.Clear

      qProblema.Add("SELECT DISTINCT G.HANDLE FROM SAM_GUIA G, SAM_GUIA_EVENTOS E WHERE G.HANDLE = E.GUIA AND G.PEG = :PEG") 'Coelho SMS: 96895

      qProblema.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      qProblema.Active = True

      qAltera.Clear

      qAltera.Add("UPDATE SAM_GUIA")
      qAltera.Add("   SET SITUACAO = (SELECT MIN(E.SITUACAO)")
      qAltera.Add("                     FROM SAM_GUIA_EVENTOS E")
      qAltera.Add("                    WHERE E.GUIA = :HANDLEGUIA AND E.COPIAEVENTOORIGINAL <> 'S' )") 'Coelho SMS: 96895
      qAltera.Add(" WHERE HANDLE = :HANDLEGUIA")

      While Not qProblema.EOF
        qAltera.ParamByName("HANDLEGUIA").Value = qProblema.FieldByName("HANDLE").AsInteger

        qAltera.ExecSQL
        qProblema.Next
      Wend

      Set qProblema = Nothing
      Set qAltera = Nothing
    End If

    AtualizarCarga(False)
  End If

  If VisibleMode Then
    If qFaseAtual <> CurrentQuery.FieldByName("SITUACAO").AsString Then
      RefreshNodesWithTableInterfacePEG("SAM_PEG")
    End If
  End If

  Set qPreencheServicos = Nothing
  Set qVerificaConsiderarSp = Nothing
  Set qServicosPrestador = Nothing
  Set Interface =Nothing

End Sub


Public Sub IMPORTARBENNER_OnClick()
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem,vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
    If CurrentQuery.State <> 1 Then
	  bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
	  Exit Sub
    End If

    Dim Interface As Object
    Set Interface = CreateBennerObject("SAMPEGDIGIT.digitacao")

    Interface.Botaoimportar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("RECEBEDOR").AsInteger, 1)

    Set Interface =Nothing

    If VisibleMode Then
      SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
    End If
  Else
	bsShowMessage("Comando válido somente na fase de digitação", "I")
  End If
End Sub


Public Sub IMPORTARSUS_OnClick()
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "2" Or _
     CurrentQuery.FieldByName("SITUACAO").AsString = "3" Or  _
     CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
	bsShowMessage("Comando válido somente na fase de digitação", "I")
	Exit Sub
  End If

  Dim SUS As Object
  Set SUS = NewQuery

  SUS.Active = False

  SUS.Clear

  SUS.Add("SELECT PRESTADORSUS FROM SAM_PARAMETROSPRESTADOR")

  SUS.Active = True

  If CurrentQuery.FieldByName("RECEBEDOR").AsInteger <>SUS.FieldByName("PRESTADORSUS").AsInteger Then
    bsShowMessage("O Recebedor não está definido como Prestador SUS!", "E")
    Exit Sub
  Else
    Dim Interface As Object
	Set Interface = CreateBennerObject("SAMPEGDIGIT.DIGITACAO")

	Interface.INICIALIZAR(CurrentSystem)
	Interface.BOTAOIMPORTAR(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("RECEBEDOR").AsInteger, 2)
	Interface.FINALIZAR
  End If
  Set SUS = Nothing
End Sub

Public Sub REPROCESSARPEG_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode reprocessar", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode reprocessar", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode reprocessar", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean
  Dim Obj As Object
  Dim Interface As Object
  Dim viRetorno As Long
  Dim vsMensagemErro As String
  Dim vvContainer As CSDContainer
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  'Aqui Reprocessa o PEG
  If (WebMode) Then
    ReprocessarPEGWeb
  Else
    Set Interface = CreateBennerObject("BSINTERFACE0046.ROTINAS")
    Interface.ReprocessarPEG(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing
  End If

  If (Not WebMode) Then
    Dim qSituacao As Object
    Set qSituacao = NewQuery

    qSituacao.Clear

    qSituacao.Add("SELECT SITUACAO FROM SAM_PEG WHERE HANDLE = :HANDLE")

    qSituacao.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
    qSituacao.Active = True

    If vSituacaoAnteriorPeg <> qSituacao.FieldByName("SITUACAO").AsString Then
      AtualizarCarga(False)
    Else
      SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
    End If
    Set qSituacao = Nothing
  End If

End Sub


Public Sub REVISAREVENTOS_OnClick()
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode revisar", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode revisar", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode revisar", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
	bsShowMessage("Comando válido somente nas fases de conferência e pronto", "I")
	Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("BSPro000.Rotinas")

  Interface.RevisarPEG(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)

  Set Interface = Nothing
  CHECATETOREEMBOLSO
End Sub


Public Sub BOTAOVERALERTAS_OnClick()
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  Dim Aux As Boolean

  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode ver alertas", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode ver alertas", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode ver alertas", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If


  Dim Interface As Object
  Set Interface = CreateBennerObject("BSINTERFACE0047.ROTINAS")

  Interface.VerAlertasPEG(CurrentSystem, CurrentQuery.FieldByName("handle").AsInteger)

  Set Interface = Nothing

  If VisibleMode Then
    SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If
End Sub


Public Sub CARTAREMESSA_OnExit()
  If ((CurrentQuery.State = 2) Or (CurrentQuery.State = 3)) Then
	Dim Valor As Long

	InicializaSamParametrosProcContas("BUSCARPEGEXTERNO, NUMERACAOPEG, SUGEREPEGCARTAREMESSA, SEMPREHERDARPEGDACARTAREMESSA")
	If ((qSamParametrosProcContas.FieldByName("BUSCARPEGEXTERNO").AsString = "S") And (CurrentQuery.FieldByName("PEG").IsNull)) Then
	  Dim qBuscaPegExterno As Object
	  Set qBuscaPegExterno = NewQuery

	  qBuscaPegExterno.Active = False
	  qBuscaPegExterno.Clear
	  qBuscaPegExterno.Add("SELECT HANDLE, 						")
	  qBuscaPegExterno.Add("	   CARTAREMESSA, 				")
	  qBuscaPegExterno.Add("	   PEG, 						")
	  qBuscaPegExterno.Add("	   RECEBEDOR, 					")
	  qBuscaPegExterno.Add("	   QTDGUIA, 					")
	  qBuscaPegExterno.Add("	   TOTALPAGARINFORMADO 			")
	  qBuscaPegExterno.Add("  FROM SAM_PEGEXTERNO 				")
	  qBuscaPegExterno.Add(" WHERE CARTAREMESSA = :pCARTAREMESSA")
	  qBuscaPegExterno.ParamByName("pCARTAREMESSA").AsFloat = CurrentQuery.FieldByName("CARTAREMESSA").AsFloat
	  qBuscaPegExterno.Active = True

	  If (Not qBuscaPegExterno.FieldByName("HANDLE").IsNull) Then ' Encontrou um registro que contém a mesma carta remessa informada
		CurrentQuery.FieldByName("PEG").AsFloat					  = qBuscaPegExterno.FieldByName("PEG").AsFloat
		CurrentQuery.FieldByName("RECEBEDOR").AsInteger			  = qBuscaPegExterno.FieldByName("RECEBEDOR").AsInteger
		CurrentQuery.FieldByName("QTDGUIA").AsInteger			  = qBuscaPegExterno.FieldByName("QTDGUIA").AsInteger
		CurrentQuery.FieldByName("TOTALPAGARINFORMADO").AsInteger = qBuscaPegExterno.FieldByName("TOTALPAGARINFORMADO").AsInteger
		CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime	  = CurrentSystem.ServerDate
	  Else
		If (Not CurrentQuery.FieldByName("CARTAREMESSA").IsNull) Then
		  If (qSamParametrosProcContas.FieldByName("NUMERACAOPEG").AsString = "M") Then
			If (qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean = True) Then
			  If (CurrentQuery.State = 3) Or _
			     ((CurrentQuery.State = 2) And _
			  	  (qSamParametrosProcContas.FieldByName("SEMPREHERDARPEGDACARTAREMESSA").AsBoolean)) Then
				CurrentQuery.FieldByName("PEG").AsFloat = CurrentQuery.FieldByName("CARTAREMESSA").AsFloat
			  End If
			End If
		  End If
		End If
	  End If

	  qBuscaPegExterno.Active = False

	  Set qBuscaPegExterno = Nothing
	Else
	  If (Not CurrentQuery.FieldByName("CARTAREMESSA").IsNull) Then
		If (qSamParametrosProcContas.FieldByName("NUMERACAOPEG").AsString = "M") Then
		  If (qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean = True) Then
			If (CurrentQuery.State = 3) Or _
			   ((CurrentQuery.State = 2) And _
			    (qSamParametrosProcContas.FieldByName("SEMPREHERDARPEGDACARTAREMESSA").AsBoolean)) Then
			  CurrentQuery.FieldByName("PEG").AsFloat = CurrentQuery.FieldByName("CARTAREMESSA").AsFloat
			End If
		  End If
		End If
	  End If
	End If
	FinalizaSamParametrosProcContas


    If Not ValidaNumeroCartaRemessa Then
      CARTAREMESSA.SetFocus
    End If
  End If
End Sub


Public Sub CONFEVENTO_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode conferir eventos", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode conferir eventos", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode conferir eventos", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim agrupadorFechado As Boolean
  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
    BsShowMessage("Não é permitida a alteração dos eventos do PEG que está ligado à registro de pagamento fechado.","E")
    Exit Sub
  End If

  Dim qBusca As Object
  Set qBusca = NewQuery

  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  qBusca.Active = False

  qBusca.Clear

  qBusca.Add("SELECT COUNT(E.HANDLE) QTDE					  ")
  qBusca.Add("  FROM SAM_GUIA_EVENTOS E,					  ")
  qBusca.Add("		 SAM_GUIA G                               ")
  qBusca.Add(" WHERE G.PEG    = :PEG")
  qBusca.Add("	 AND G.HANDLE = E.GUIA  					  ")
  qBusca.Add("   AND E.COPIAEVENTOORIGINAL <> 'S'             ")
  qBusca.Add("   AND (EXISTS(SELECT GL.HANDLE				  ")
  qBusca.Add("				   FROM SAM_GUIA_EVENTOS_GLOSA GL ")
  qBusca.Add("				  WHERE GL.GUIAEVENTO = E.HANDLE  ")
  qBusca.Add("					AND GL.GLOSAREVISADA = 'N')	  ")
  qBusca.Add("    OR  EXISTS(SELECT N.HANDLE				  ")
  qBusca.Add("				   FROM SAM_GUIA_EVENTOS_NEGACAO N")
  qBusca.Add("				  WHERE N.GUIAEVENTO = E.HANDLE	  ")
  qBusca.Add("					AND N.NEGACAOREVISADA = 'N')  ")
  qBusca.Add("		 )										  ")

  qBusca.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qBusca.Active = True

  If qBusca.FieldByName("QTDE").AsInteger <= 0 Then
    bsShowMessage("Não existem eventos a serem conferidos!", "I")
    Set qBusca = Nothing
  End If

  If (Not UtilizaTelaConferenciaIntegrada) Then
    Dim DLL49 As Object
    Set DLL49 = CreateBennerObject("BSINTERFACE0049.Conferencia")
    DLL49.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "EVENTO", "TREEVIEW_PEG", -1)
    Set DLL49 = Nothing
  Else
    Dim DLL64 As Object
    Set DLL64 = CreateBennerObject("BSINTERFACE0064.CONFERENCIA")
    DLL64.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "EVENTO", "T", 0)
    Set DLL64 = Nothing
  End If

  AtualizarCarga(False)

End Sub

Public Function UtilizaTelaConferenciaIntegrada As Boolean
   Dim sql As BPesquisa
   Set sql = NewQuery
   sql.Add("SELECT TELADECONFERENCIAINTEGRADA FROM SAM_PARAMETROSPROCCONTAS")
   sql.Active=True

   UtilizaTelaConferenciaIntegrada = (sql.FieldByName("TELADECONFERENCIAINTEGRADA").AsString = "S")
   Set sql = Nothing
End Function

Public Sub CONFGUIA_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode conferir guias", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode conferir guias", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode conferir guias", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim agrupadorFechado As Boolean
  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
    BsShowMessage("Não é permitida a alteração dos eventos do PEG que está ligado à registro de pagamento fechado.","E")
    Exit Sub
  End If

  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If


  If (Not UtilizaTelaConferenciaIntegrada) Then
    Dim DLL49 As Object
    Set DLL49 = CreateBennerObject("BSINTERFACE0049.Conferencia")
    DLL49.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "GUIA", "TREEVIEW_PEG", -1)
    Set DLL49 = Nothing
  Else
    Dim DLL64 As Object
    Set DLL64 = CreateBennerObject("BSINTERFACE0064.CONFERENCIA")
    DLL64.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "GUIA", "T", 0)
    Set DLL64 = Nothing
  End If

End Sub


Public Sub CONFPEG_OnClick()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não pode conferir peg", "I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    bsShowMessage("PEG devolvido não pode conferir peg", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "8" Then
    bsShowMessage("PEG cancelado não pode conferir peg", "I")
    Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
    bsShowMessage("O registro está em edição! Por favor confirme ou cancele as alterações", "I")
    Exit Sub
  End If

  Dim agrupadorFechado As Boolean
  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
    BsShowMessage("Não é permitida a alteração dos eventos do PEG que está ligado à registro de pagamento fechado.","E")
    Exit Sub
  End If

  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  If (Not UtilizaTelaConferenciaIntegrada) Then
    Dim DLL49 As Object
    Set DLL49 = CreateBennerObject("BSINTERFACE0049.Conferencia")
    DLL49.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "PEG", "TREEVIEW_PEG", -1)
    Set DLL49 = Nothing
  Else
    Dim DLL64 As Object
    Set DLL64 = CreateBennerObject("BSINTERFACE0064.CONFERENCIA")
    DLL64.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "PEG", "T", 0)
    Set DLL64 = Nothing
  End If

  AtualizarCarga(False)
End Sub


Public Sub DATAADIANTAMENTO_OnExit()
  If DATAADIANTAMENTO.ReadOnly = False Then
    gchangeDataAdi = False
    VerificaAdiantamento False, True
  End If
End Sub


Public Sub DATAPAGAMENTO_OnExit()
  vPodeSAlvarPegDataPagamentoCorreta = True
  vVerificaDataWeb = False

  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    If (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime) Or _
       (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < ServerDate) Then
      bsShowMessage("A data de pagamento deve ser maior que a de recebimento e maior ou igual a de hoje", "E")
      vVerificaDataWeb = True
      Exit Sub
    End If
  End If

  ACHOU = True

  If CurrentQuery.State = 1 Then
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
    Exit Sub
  End If

  InicializaSamParametrosProcContas("UTILIZACALENDARIODIARIO, CALENDARIOEXCECAO")
  vgUtilizaCalendarioDiario = qSamParametrosProcContas.FieldByName("UTILIZACALENDARIODIARIO").AsString
  vgCalendarioExcecao = qSamParametrosProcContas.FieldByName("CALENDARIOEXCECAO").AsString
  FinalizaSamParametrosProcContas

  If OLDDATAPAGAMENTO <> CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime Then
    OLDDATAPAGAMENTO = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime
    vMudouDataPagamentoPEG = True

    ARRUMAROTULOS
    CalculaAdiantamento

    If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
      If Not CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
        Dim Interface As Object
        Set Interface = CreateBennerObject("SAMCALENDARIOPGTO.ROTINAS")

        Interface.INICIALIZAR(CurrentSystem)

        If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime <> Interface.DIAUTILANTERIOR(CurrentSystem, CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime) Then
          bsShowMessage("Entre com um dia útil para a Data de Pagamento", "E")

          DATAPAGAMENTO.SetFocus

          ACHOU = False
          Interface.FINALIZAR
          Set Interface = Nothing
          OLDDATAPAGAMENTO = 0
          vVerificaDataWeb = True

          Exit Sub
        End If
        'verificar se o processamento do dia ja foi efetuado
        Interface.FINALIZAR

        Dim q1 As Object
        Set q1 = NewQuery
        q1.Clear

        If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
          q1.Add("SELECT DATAPROCESSAMENTO, DATAFECHAMENTO")
          q1.Add("  FROM SAM_PAGAMENTO")
          q1.Add(" WHERE DATAPAGAMENTO = :DATAPGTO")
          q1.Add(" ORDER BY DATAFECHAMENTO")
        Else
          q1.Add("SELECT DATAPROCESSAMENTO, DATAFECHAMENTO")
          q1.Add("  FROM SAM_CALENDARIOREEMBOLSO")
          q1.Add(" WHERE DATAPAGAMENTO = :DATAPGTO")
        End If

        q1.ParamByName("DATAPGTO").Value = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime
        q1.Active = True

        If q1.EOF Then
          If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
            If vgUtilizaCalendarioDiario ="N" Then
              bsShowMessage("Data de pagamento não permitida - data não cadastrada no calendário geral.", "E")
              CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0
              OLDDATAPAGAMENTO = 0
              ACHOU = False

              DATAPAGAMENTO.SetFocus
              vVerificaDataWeb = True

              Exit Sub
            Else
              If CurrentQuery.FieldByName("NUMEROPAGAMENTO").AsInteger <= 0 Then
                CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0
              End If

              Exit Sub
            End If
          Else
            CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0

            If vgUtilizaCalendarioDiario ="N" Then
              VERIFICAPAG("I")
            End If
          End If
        Else
          If q1.FieldByName("DATAFECHAMENTO").IsNull Then
            If Interface.PegarNumPagamento(CurrentSystem, CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime, _
            							   CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger, _
            							   CurrentQuery.FieldByName("TIPOPEG").AsInteger, AUXNUMPAG) = True Then
              CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = AUXNUMPAG
            Else
              If CurrentQuery.FieldByName("NUMEROPAGAMENTO").AsInteger <= 0 Then
                CurrentQuery.FieldByName("NUMEROPAGAMENTO").Value = 0
              End If
            End If
          Else
            VERIFICAPAG("I")
            bsShowMessage("Data de Pagamento não Permitida - calendário fechado.", "E")
            vPodeSAlvarPegDataPagamentoCorreta = False

            If vgUtilizaCalendarioDiario = "N" Then
              OLDDATAPAGAMENTO = 0
            End If

			InicializaSamParametrosProcContas("UTILIZACALENDARIODIARIO")
            If qSamParametrosProcContas.FieldByName("UTILIZACALENDARIODIARIO").AsString = "N" Then
              vVerificaDataWeb = True
              ACHOU = False
            End If
			FinalizaSamParametrosProcContas

            Exit Sub
          End If
        End If

        q1.Active = False
        Set q1 = Nothing

        Set Interface = Nothing
      End If
    End If
  Else
  	vMudouDataPagamentoPEG = False
  End If
End Sub


Public Sub DATARECEBIMENTO_OnExit()
  Dim vCOMPETENCIA As Long
  If (VisibleMode Or WebMode) Then
	vCOMPETENCIA = RecordHandleOfTableInterfacePEG("SAM_COMPETPEG")
  Else
	vCOMPETENCIA = CurrentQuery.FieldByName("COMPETENCIA").AsInteger
  End If

  If vCOMPETENCIA <= 0 And CurrentQuery.State <> 1 Then
    Dim sql2 As Object
    Set sql2 = NewQuery

    sql2.Add("SELECT HANDLE, COMPETENCIA FROM SAM_COMPETPEG WHERE COMPETENCIA < :DATA ORDER BY COMPETENCIA DESC")

    sql2.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime
    sql2.Active = True

    CurrentQuery.FieldByName("COMPETENCIA").AsInteger = sql2.FieldByName("HANDLE").AsInteger

    Set sql2 = Nothing
  End If


  If CurrentQuery.State = 1 Then
	Exit Sub
  End If

  Dim testa As Date
  On Error GoTo erro

  testa = CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime

  If OLDDATARECEBIMENTO <> CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime Then
    OLDDATARECEBIMENTO = CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime
     VERIFICAPAG("I") 'VERIFICA A DATA PAGAMENTO E O NUM PAGAMENTO
    ARRUMAROTULOS
    CalculaAdiantamento
    ' Verifica se a data de recebimento do PEG está válida
    If Not VerificaDataRecebimento Then
      bsShowMessage("Data Recebimento não pode ser menor que a data atendimento de alguma guia do PEG", "I")
    End If
  End If
  erro :
End Sub


Public Sub FILIALPROCESSAMENTO_OnPopup(ShowPopup As Boolean)
  FILIALPROCESSAMENTO.LocalWhere = "HANDLE IN (SELECT F.FILIALPROCESSAMENTO " + _
  								   "			 FROM Z_GRUPOUSUARIOS A, FILIAIS F " + _
  								   "			WHERE (A.FILIALPADRAO = F.Handle) " + _
  								   "			  AND (A.Handle = " + Str(CurrentUser) + ")" + _
  								   "			UNION " + _
  								   "		   SELECT F.FILIALPROCESSAMENTO " + _
  								   "			 FROM Z_GRUPOUSUARIOS_FILIAIS A, FILIAIS F " + _
  								   "			WHERE (A.FILIAL = F.Handle) AND (A.USUARIO = " + Str(CurrentUser) + "))"
End Sub


Public Sub LOCALEXECUCAO_OnPopup(ShowPopup As Boolean)
  Dim vTipoBusca As String
  Dim vHandle As Long

  ShowPopup = False

  If (IsNumeric(LOCALEXECUCAO.Text)) Then
      vTipoBusca = "C"
  Else
      vTipoBusca = "N"
  End If

  vHandle = ProcuraPrestador(vTipoBusca, "L", LOCALEXECUCAO.Text)

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("LOCALEXECUCAO").Value = vHandle
  End If

End Sub


Public Sub BENEFICIARIO_OnExit()
  Dim sqlfp As Object
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  If CurrentQuery.State <>1 Then
    VERIFICAPAG("I") 'VERIFICA A DATA PAGAMENTO E O NUM PAGAMENTO

	If (vFilial = -1) And _
	  CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2 And _
	  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger > 0 Then
	  Set sqlfp = NewQuery

	  sqlfp.Clear

	  sqlfp.Add("SELECT FILIALCUSTO FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")

	  sqlfp.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
	  sqlfp.Active = True

	  If Not sqlfp.FieldByName("FILIALCUSTO").IsNull Then
	    vgFilial = sqlfp.FieldByName("FILIALCUSTO").AsInteger
	    vgFilialProcessamento = BuscarFilialProcessamento(CurrentSystem, sqlfp.FieldByName("FILIALCUSTO").AsInteger)

		  CurrentQuery.FieldByName("FILIAL").AsInteger = vgFilial
		  CurrentQuery.FieldByName("FILIALPROCESSAMENTO").AsInteger = vgFilialProcessamento
	  Else
	    vgFilial = 0
	  End If

	  Set sqlfp = Nothing
	End If
  End If
End Sub


Public Sub PEG_OnExit()
  If CurrentQuery.State = 1 Then
	Exit Sub
  End If

  If CurrentQuery.State = 3 Then 'inserir

    InicializaSamParametrosProcContas("BUSCARPEGEXTERNO, NUMERACAOPEG, REPETICAONUMERACAOPEG, SUGEREPEGCARTAREMESSA, SEMPREHERDARPEGDACARTAREMESSA")
	If ((qSamParametrosProcContas.FieldByName("BUSCARPEGEXTERNO").AsString = "S") And (CurrentQuery.FieldByName("CARTAREMESSA").IsNull)) Then
	  Dim qBuscaPegExterno As Object
	  Set qBuscaPegExterno = NewQuery

	  qBuscaPegExterno.Active = False
	  qBuscaPegExterno.Clear
	  qBuscaPegExterno.Add("SELECT HANDLE,			  ")
	  qBuscaPegExterno.Add("	   CARTAREMESSA,	  ")
	  qBuscaPegExterno.Add("	   PEG,				  ")
	  qBuscaPegExterno.Add("	   RECEBEDOR,		  ")
	  qBuscaPegExterno.Add("	   QTDGUIA,			  ")
	  qBuscaPegExterno.Add("	   TOTALPAGARINFORMADO")
	  qBuscaPegExterno.Add("  FROM SAM_PEGEXTERNO	  ")
	  qBuscaPegExterno.Add(" WHERE PEG = :pPEG		  ")

	  qBuscaPegExterno.ParamByName("pPEG").AsFloat = CurrentQuery.FieldByName("PEG").AsFloat
	  qBuscaPegExterno.Active = True

	  If (Not qBuscaPegExterno.FieldByName("HANDLE").IsNull) Then ' Encontrou um registro que contém a mesma carta remessa informada
		CurrentQuery.FieldByName("CARTAREMESSA").AsFloat			= qBuscaPegExterno.FieldByName("CARTAREMESSA").AsFloat
		CurrentQuery.FieldByName("RECEBEDOR").AsInteger				= qBuscaPegExterno.FieldByName("RECEBEDOR").AsInteger
		CurrentQuery.FieldByName("QTDGUIA").AsInteger				= qBuscaPegExterno.FieldByName("QTDGUIA").AsInteger
		CurrentQuery.FieldByName("TOTALPAGARINFORMADO").AsInteger	= qBuscaPegExterno.FieldByName("TOTALPAGARINFORMADO").AsInteger
		CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime		= CurrentSystem.ServerDate
	  End If

	  qBuscaPegExterno.Active = False

	  Set qBuscaPegExterno = Nothing
	End If

	Dim q1 As Object
	Dim Comp As Object
	Set q1 = NewQuery

	q1.Clear

	'buscar COMPETENCIA e PEG,de acordo com a chave única

	If qSamParametrosProcContas.FieldByName("NUMERACAOPEG").AsString = "M" Then
	  Select Case qSamParametrosProcContas.FieldByName("REPETICAONUMERACAOPEG").AsString
	  Case "2"
		q1.Add("SELECT HANDLE FROM SAM_PEG WHERE PEG=:PEG AND SITUACAO <>'9' AND SEQUENCIA = :SEG")

		q1.ParamByName("SEG").Value = CurrentQuery.FieldByName("SEQUENCIA").AsFloat
		q1.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsFloat
		q1.Active = True

		If Not q1.EOF Then
		  bsShowMessage("Já existe um PEG com este número", "E")
		  PEG.SetFocus
		End If
	  Case "3"
		q1.Add("SELECT HANDLE FROM SAM_PEG WHERE COMPETENCIA=:COMPETENCIA AND PEG=:PEG AND SITUACAO <>'9' AND SEQUENCIA = :SEG")

		q1.ParamByName("COMPETENCIA").Value = CurrentQuery.FieldByName("COMPETENCIA").AsFloat
		q1.ParamByName("SEG").Value = CurrentQuery.FieldByName("SEQUENCIA").AsFloat
		q1.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsFloat
		q1.Active = True

		If Not q1.EOF Then
		  bsShowMessage("Já existe um PEG com este número nesta competência", "E")
		  PEG.SetFocus
		End If
	  Case "4"
		q1.Add("SELECT HANDLE FROM SAM_PEG WHERE SEQUENCIA = :SEG AND PEG=:PEG AND SITUACAO <>'9' AND FILIAL=:FILIAL")

		q1.ParamByName("SEG").Value = CurrentQuery.FieldByName("SEQUENCIA").AsFloat
		q1.ParamByName("FILIAL").Value = CurrentQuery.FieldByName("FILIAL").AsFloat
		q1.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsFloat
		q1.Active = True

		If Not q1.EOF Then
		  bsShowMessage("Já existe um PEG com este número nesta filial", "E")
		  PEG.SetFocus
		End If
	  End Select
	  q1.Active = False
	  Set q1 = Nothing
	End If

    If CurrentQuery.FieldByName("CARTAREMESSA").IsNull Then
	  If (qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean) Or (qSamParametrosProcContas.FieldByName("SEMPREHERDARPEGDACARTAREMESSA").AsBoolean) Then
  	    CurrentQuery.FieldByName("CARTAREMESSA").Value = CurrentQuery.FieldByName("PEG").Value
      End If
    End If
	FinalizaSamParametrosProcContas
  End If
End Sub


Public Sub QTDGUIA_OnExit()
  If CurrentQuery.State <> 1 Then
    If CurrentQuery.FieldByName("QTDGUIAINFORMADA").IsNull Then
	  CurrentQuery.FieldByName("QTDGUIAINFORMADA").Value = CurrentQuery.FieldByName("QTDGUIA").AsInteger
    End If
  End If
End Sub


Public Sub RECEBEDOR_OnExit()

  Dim sql As Object
  Dim vJahVerificouPag As Boolean
  vJahVerificouPag = False
  VerificaReciboNF

  Set sql = NewQuery

  If CurrentQuery.State = 1 Then 'browse
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then

	If OLDRECEBEDOR <> CurrentQuery.FieldByName("RECEBEDOR").AsInteger Or CurrentQuery.FieldByName("PERCENTUALREDUCAOINSS").IsNull Then
		AtualizarDescontoINSS
	End If


    sql.Add("SELECT LOCALEXECUCAO, FILIALPADRAO FROM SAM_PRESTADOR WHERE HANDLE = " + CurrentQuery.FieldByName("RECEBEDOR").AsString)

    sql.Active = True

    If sql.FieldByName("LOCALEXECUCAO").AsString = "S" Then
      CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    End If
	Dim vFilial As Long
    If (VisibleMode Or WebMode) Then
	  vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
    Else
	  vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
    End If
    If (vFilial <= 0) Then
      vgFilial = sql.FieldByName("FILIALPADRAO").AsInteger

      sql.Clear

      sql.Add("SELECT FILIALPROCESSAMENTO FROM FILIAIS WHERE HANDLE=" + Str(vgFilial))

      sql.Active = True

      vgFilialProcessamento = sql.FieldByName("FILIALPROCESSAMENTO").AsInteger

      If vgFilial = 0 Then
        bsShowMessage("O recebedor está sem filial padrão", "I")
      Else
        If vgFilialProcessamento > 0 Then
          CurrentQuery.FieldByName("FILIALPROCESSAMENTO").AsInteger = vgFilialProcessamento
        Else
          bsShowMessage("Filial padrão sem filial de processamento", "I")
        End If

        CurrentQuery.FieldByName("FILIAL").AsInteger = vgFilial

        FILIAL.ReadOnly = AtribuirReadOnly(True)
        FILIALPROCESSAMENTO.ReadOnly = AtribuirReadOnly(True)
      End If
    End If
  End If
  Set sql = Nothing

  On Error GoTo erro
    If OLDRECEBEDOR <> CurrentQuery.FieldByName("RECEBEDOR").AsInteger Then
      OLDRECEBEDOR = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
      VERIFICAPAG("I") 'VERIFICA A DATA PAGAMENTO E O NUM PAGAMENTO
      vJahVerificouPag = True
    End If

  Erro :
  InicializaSamParametrosProcContas("PERMITIRIMPOSTOSNOPEG")
  If(qSamParametrosProcContas.FieldByName("PERMITIRIMPOSTOSNOPEG").AsString = "S")Then
    Dim q1 As Object
    Set q1 = NewQuery

    q1.Add("SELECT CD.CODIGORETENCAO")
    q1.Add("FROM SAM_PRESTADOR_IRRF P, SFN_CODIGODIRF CD")
    q1.Add("WHERE P.PRESTADOR = :PHANDLE")
    q1.Add("      AND CD.HANDLE = P.CODIGODIRF")

    q1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    q1.Active = True

    If Not q1.FieldByName("CODIGORETENCAO").IsNull Then
      CurrentQuery.FieldByName("CODIGORECEITAIRRF").AsInteger = q1.FieldByName("CODIGORETENCAO").AsInteger
    End If

    q1.Active = False

    q1.Clear

    q1.Add("SELECT I.DESCRICAO")
    q1.Add("FROM SAM_PRESTADOR P, SFN_ISS I")
    q1.Add("WHERE P.HANDLE = :PHANDLE")
    q1.Add("      AND I.HANDLE = P.ISS")

    q1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    q1.Active = True

    If Not q1.FieldByName("DESCRICAO").IsNull Then
      CurrentQuery.FieldByName("CODIGOTRIBUTACAOISS").AsString = q1.FieldByName("DESCRICAO").AsString
    End If

    q1.Active = False

    q1.Clear

    q1.Add("SELECT CD.CODIGORETENCAO")
    q1.Add("FROM SAM_PRESTADOR_IRRF P, SFN_CODIGODIRF CD")
    q1.Add("WHERE P.PRESTADOR = :PHANDLE")
    q1.Add("      AND CD.HANDLE = P.CSLLCODIGODIRF")

    q1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    q1.Active = True

    If Not q1.FieldByName("CODIGORETENCAO").IsNull Then
      CurrentQuery.FieldByName("CODIGORECEITACSLL").AsInteger = q1.FieldByName("CODIGORETENCAO").AsInteger
    End If

    q1.Active = False
    q1.Clear

    q1.Add("SELECT CD.CODIGORETENCAO")
    q1.Add("FROM SAM_PRESTADOR_IRRF P, SFN_CODIGODIRF CD")
    q1.Add("WHERE P.PRESTADOR = :PHANDLE")
    q1.Add("      AND CD.HANDLE = P.COFINSCODIGODIRF")

    q1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    q1.Active = True

    If Not q1.FieldByName("CODIGORETENCAO").IsNull Then
      CurrentQuery.FieldByName("CODIGORECEITACOFINS").AsInteger = q1.FieldByName("CODIGORETENCAO").AsInteger
    End If

    q1.Active = False
    q1.Clear

    q1.Add("SELECT CD.CODIGORETENCAO")
    q1.Add("FROM SAM_PRESTADOR_IRRF P, SFN_CODIGODIRF CD")
    q1.Add("WHERE P.PRESTADOR = :PHANDLE")
    q1.Add("      AND CD.HANDLE = P.PISPASEPCODIGODIRF")

    q1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
    q1.Active = True

    If Not q1.FieldByName("CODIGORETENCAO").IsNull Then
      CurrentQuery.FieldByName("CODIGORECEITAPISPASEP").AsInteger = q1.FieldByName("CODIGORETENCAO").AsInteger
    End If

    Set q1 = Nothing
  End If
  FinalizaSamParametrosProcContas

  If ((CurrentQuery.State = 3) Or (CurrentQuery.State = 2)) Then
    If EhPrestadorFixo(CurrentQuery.FieldByName("RECEBEDOR").AsInteger) Then
      If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
        VALORPAGAMENTORATEIO.ReadOnly = False
      End If
    Else
      CurrentQuery.FieldByName("VALORPAGAMENTORATEIO").Clear
      VALORPAGAMENTORATEIO.ReadOnly = True
    End If
  End If

  If Not vJahVerificouPag Then
    VERIFICAPAG("I")
  End If
End Sub


Public Sub RECEBEDOR_OnPopup(ShowPopup As Boolean)
  Dim vTipoBusca As String
  Dim vHandle As Long

  ShowPopup = False

  If (IsNumeric(RECEBEDOR.Text)) Then
      vTipoBusca = "C"
  Else
      vTipoBusca = "N"
  End If

  vHandle = ProcuraPrestador(vTipoBusca, "R", RECEBEDOR.Text)

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("RECEBEDOR").Value = vHandle
  End If

End Sub


Public Sub TABLE_AfterCancel()
  InicializaSamParametrosProcContas("NUMERACAOPEG")
  If (NodeInternalCode <> 10) And (qSamParametrosProcContas.FieldByName("NUMERACAOPEG").AsString = "A") Then
    'read only para o campo peg
    PEG.ReadOnly = AtribuirReadOnly(True)
  Else
    PEG.ReadOnly = AtribuirReadOnly(False)
  End If
  FinalizaSamParametrosProcContas

End Sub


Public Sub TABLE_AfterCommitted()

  Dim q1 As Object
  Dim Interface As Object
  Dim qVerificaCT As Object
  Dim qVerificaCT2 As Object
  Dim interfaceBspro006 As Object
  Set interfaceBspro006 = CreateBennerObject("BSPRO006.ROTINAS")


  Set q1 = NewQuery
  q1.Clear
  q1.Add("UPDATE SAM_GUIA SET NUMEROPAGAMENTO=:NUMEROPAGAMENTO WHERE PEG=:PEG")
  q1.ParamByName("NUMEROPAGAMENTO").Value =CurrentQuery.FieldByName("NUMEROPAGAMENTO").AsInteger
  q1.ParamByName("PEG").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  q1.ExecSQL
  Set q1 = Nothing

  If Not WebMode Then
  	If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
	    Set Interface = CreateBennerObject("SAMPEG.PROCESSAR")
	    Interface.alertasRecebedor(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	    Set Interface = Nothing
  	End If
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
    'Atualiza o registro na SAM_PEGEXTERNO caso os dados foram importados dessa tabela
	InicializaSamParametrosProcContas("DIGITACAOAUTOMATICA, BUSCARPEGEXTERNO")
    If (qSamParametrosProcContas.FieldByName("BUSCARPEGEXTERNO").AsString = "S") Then

	  Dim qUpdate As Object
	  Dim qFilial As Object
	  Set qUpdate = NewQuery
	  Set qFilial = NewQuery

	  qFilial.Active = False
	  qFilial.Clear
	  qFilial.Add("SELECT F.HANDLE FROM Z_GRUPOUSUARIOS U, FILIAIS F")
	  qFilial.Add(" WHERE U.FILIALPADRAO = F.HANDLE AND F.FILIALPROCESSAMENTO = F.HANDLE AND U.HANDLE = :pUSUARIO")
	  qFilial.ParamByName("pUSUARIO").AsInteger = CurrentUser
	  qFilial.Active = True

	  qUpdate.Active = False
	  qUpdate.Clear
	  qUpdate.Add("UPDATE SAM_PEGEXTERNO SET DATARECEBIMENTO = :pDATARECEBIMENTO")

	  If (Not qFilial.EOF) Then
	    qUpdate.Add(", DATARECEBIMENTOFILIALPROC = :pDATARECEBIMENTO")
	  End If

	  qUpdate.Add("    WHERE CARTAREMESSA = :pCARTAREMESSA AND PEG = :pPEG")
	  qUpdate.ParamByName("pCARTAREMESSA").AsFloat		= CurrentQuery.FieldByName("CARTAREMESSA").AsFloat
	  qUpdate.ParamByName("pPEG").AsFloat					= CurrentQuery.FieldByName("PEG").AsFloat
	  qUpdate.ParamByName("pDATARECEBIMENTO").AsDateTime	= CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime
	  qUpdate.ExecSQL
	  qUpdate.Active = False
	  qFilial.Active = False

	  Set qUpdate = Nothing
	  Set qFilial = Nothing
    End If

    If qSamParametrosProcContas.FieldByName("DIGITACAOAUTOMATICA").AsString = "S" Then

      If (CurrentQuery.FieldByName("CREDITOCONTATERCEIROS").AsString = "S") And (CurrentQuery.FieldByName("BANCO").IsNull) Then
        interfaceBspro006.ContaTerceiro(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

        Set qVerificaCT = NewQuery
	    qVerificaCT.Clear
        qVerificaCT.Add("SELECT BANCO,              ")
        qVerificaCT.Add("       AGENCIA,            ")
        qVerificaCT.Add("       CONTACORRENTENUMERO,")
        qVerificaCT.Add("       CONTACORRENTEDV,    ")
        qVerificaCT.Add("       CONTACORRENTENOME,  ")
        qVerificaCT.Add("       CONTACORRENTECPFCNPJ")
        qVerificaCT.Add("  FROM SAM_PEG             ")
        qVerificaCT.Add(" WHERE HANDLE =:HANDLEPEG  ")
        qVerificaCT.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qVerificaCT.Active = True

        If qVerificaCT.FieldByName("BANCO").IsNull Then
          bsShowMessage("Conta de terceiro não informada. O campo 'Crédito em conta de terceiro' será desmarcado.", "I")
          CurrentQuery.Edit
          CurrentQuery.FieldByName("CREDITOCONTATERCEIROS").AsString = "N"
          CurrentQuery.Post
        End If

        Set qVerificaCT = Nothing
      End If

	  If Not WebMode Then

		If (Not PermitirDigitacaoPeg()) Then

	        If PermitePeloGrupoSeguranca("SAM_PEG", "BOTAODIGITAR") Then
		        Dim Interface3 As Object
		        Set Interface3 = CreateBennerObject("BSPRO006.ROTINAS")
			    Interface3.Digitar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
		        Set Interface3 = Nothing
		   	Else
		   	  bsShowMessage("Usuário não pode digitar guias/eventos devido ao grupo de segurança", "I")
		    End If
		End If
      End If
    End If
	FinalizaSamParametrosProcContas

  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "2" Or CurrentQuery.FieldByName("SITUACAO").AsString = "3" Or CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
    Dim vPeg As Long
    Dim SAMPEGF As Object
    Set SAMPEGF = CreateBennerObject("sampeg.processar")
    vPeg = CurrentQuery.FieldByName("HANDLE").AsInteger

    InicializaSamParametrosProcContas("OBRIGATITULARREEMBOLSO")
    If qSamParametrosProcContas.FieldByName("OBRIGATITULARREEMBOLSO").AsString = "S" Then
      If (vOLDBeneficiarioTitular <> CurrentQuery.FieldByName("BENEFICIARIO").AsInteger) Then
        If (WebMode) Then
          ReprocessarPEGWeb
        Else
          SAMPEGF.VerificarEventosPeg(CurrentSystem,vPeg)
          SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
          RefreshNodesWithTableInterfacePEG("SAM_PEG")
        End If
      End If
	End If
	FinalizaSamParametrosProcContas

    If (OLDRECEBEDOR <> CurrentQuery.FieldByName("RECEBEDOR").AsInteger) Or _
      (OLDLOCALEXECUCAO <> CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger) And _
      (vbEditando = True) And (CurrentQuery.FieldByName("SITUACAO").AsString > "1") Then
      'Atualizar todos os recebedores dos eventos da guia, cujo o campo recebedor não esteja no modelo da guia.
      Dim qGuia As Object
      Dim qUpdateGuia As Object
      Dim qUpdateGuiaEvento As Object

      Set qGuia = NewQuery
      Set qUpdateGuia = NewQuery
      Set qUpdateGuiaEvento = NewQuery

      qGuia.Clear

      qGuia.Add("SELECT GUIA.HANDLE")
      qGuia.Add("  FROM SAM_GUIA GUIA")
      qGuia.Add(" WHERE GUIA.PEG = :PEG")
      qGuia.Add("   AND NOT EXISTS (SELECT 1")
      qGuia.Add("                     FROM SAM_TIPOGUIA_MDGUIA M")
      qGuia.Add("                     JOIN SAM_TIPOGUIA_MDGUIA_EVENTO E ON (M.HANDLE = E.MODELOGUIA)")
      qGuia.Add("                     JOIN SIS_MODELOGUIA_CAMPOS S ON (S.HANDLE = E.SISCAMPO)")
      qGuia.Add("                     JOIN SAM_GUIA G ON (G.MODELOGUIA = M.HANDLE)")
      qGuia.Add("                    WHERE G.HANDLE = GUIA.HANDLE")
      qGuia.Add("                      AND S.ZCAMPO = 'RECEBEDOR')")

      qGuia.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qGuia.Active = True

      qUpdateGuia.Clear

      qUpdateGuia.Add("UPDATE SAM_GUIA SET RECEBEDOR = :RECEBEDOR WHERE HANDLE = :HANDLE")

      qUpdateGuiaEvento.Clear

      qUpdateGuiaEvento.Add("UPDATE SAM_GUIA_EVENTOS SET RECEBEDOR = :RECEBEDOR WHERE GUIA = :HANDLE AND COPIAEVENTOORIGINAL <> 'S' ")

      While Not qGuia.EOF
        qUpdateGuia.ParamByName("HANDLE").AsInteger = qGuia.FieldByName("HANDLE").AsInteger
        qUpdateGuia.ParamByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger

        qUpdateGuia.ExecSQL

        qUpdateGuiaEvento.ParamByName("HANDLE").AsInteger = qGuia.FieldByName("HANDLE").AsInteger
        qUpdateGuiaEvento.ParamByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger

        qUpdateGuiaEvento.ExecSQL
        qGuia.Next
      Wend

      'Aqui deve-se reprocessar o Peg.
      If WebMode Then
        ReprocessarPEGWeb
      Else
	    SAMPEGF.VerificarEventosPeg(CurrentSystem,vPeg)
        SelectNodeInterfacePEG(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
        RefreshNodesWithTableInterfacePEG("SAM_PEG")
      End If
    End If
    Set SAMPEGF = Nothing
  End If

  vbEditando = False

  Set Interface =Nothing

  OLDLOCALEXECUCAO = 0
  OLDRECEBEDOR = 0

  InicializaSamParametrosProcContas("NUMERACAOPEG")
  If (NodeInternalCode <> 10) And (qSamParametrosProcContas.FieldByName("NUMERACAOPEG").AsString = "A") Then
    'habilita o campo
     PEG.ReadOnly = AtribuirReadOnly(True)
  Else
     PEG.ReadOnly = AtribuirReadOnly(False)
  End If
  FinalizaSamParametrosProcContas

  Set interfaceBspro006 = Nothing

  If (gTabOrigemRecursoPEG <> CurrentQuery.FieldByName("TABORIGEMRECURSOPEG").AsInteger) Then

    Dim vsMensagemErro As String
    Dim Obj As Object
    Dim viRet As Long

  	Set Obj = CreateBennerObject("SAMPEG.ReclassificarPEG")
	viRet = Obj.Exec(CurrentSystem, _
		             CurrentQuery.FieldByName("HANDLE").AsInteger, _
		             0, _
		             vsMensagemErro)

	Set Obj = Nothing

    If viRet = 0 Then
      bsShowMessage("Realizado reclassificação das Guias por troca de Origem do Recurso com sucesso!", "I")
    Else
      bsShowMessage("Problema com reclassificação das Guias por troca de Origem do Recurso! Executar o comando de Reclassificação de Guias para maiores detalhes", "E")
    End If

  End If

End Sub


Public Sub TABLE_AfterEdit()
  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Add("SELECT COUNT(*) QTDGUIAS")
  qSQL.Add("FROM SAM_GUIA")
  qSQL.Add("WHERE PEG = :HPEG")
  qSQL.ParamByName("HPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSQL.Active = True

  If qSQL.FieldByName("QTDGUIAS").AsInteger > 0 Then
    TIPODEGUIA.ReadOnly = True
  Else
    TIPODEGUIA.ReadOnly = False
  End If

  Set qSQL = Nothing
End Sub


Public Sub TABLE_AfterInsert()
  Dim VALOR As Long
  Dim vRegime As Integer
  Dim qVersaoTiss As Object
  Set qVersaoTiss = NewQuery
  qVersaoTiss.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
  qVersaoTiss.Active = True
  CurrentQuery.FieldByName("VERSAOTISS").AsInteger = qVersaoTiss.FieldByName("HANDLE").AsInteger
  Set qVersaoTiss = Nothing
  TIPODEGUIA.ReadOnly = False

  InicializaSamParametrosProcContas("NUMERACAOPEG")
  If qSamParametrosProcContas.FieldByName("NUMERACAOPEG").AsString = "A" Then
    'É necessário que o contador da SAMPEG seja chamado por procedure, pois pegs serao criados via stored procedure. Logo todos os contadores da sam_peg devem ser chamados por procedure
    Dim vDLL As Object
    Dim vResult As String
    Dim vFilial As Long
    Dim param As String

    If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
    Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
    End If

    param = "p_NomeContador;SAM_PEG;S;I|P_CHAVE;"+Str(vFilial)+";I;I|P_INCREMENTA;1;I;I|P_CONTADOR;0;I;O|"

    Set vDLL = CreateBennerObject("SAMPROCEDURE.UIPROCEDURE")

    vDLL.ExecProc(CurrentSystem, "BS_NEWCONTER", param, vResult)
    vResult = Mid(vResult, InStr(1, vResult, "=") + 1, Len(vResult) - InStr(1, vResult, "=") - 1)

    Set vDLL = Nothing

    VALOR = CLng(vResult)

    CurrentQuery.FieldByName("PEG").Value = VALOR

    InicializaSamParametrosProcContas("SUGEREPEGCARTAREMESSA, SEMPREHERDARPEGDACARTAREMESSA")
    If (qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean) Or (qSamParametrosProcContas.FieldByName("SEMPREHERDARPEGDACARTAREMESSA").AsBoolean) Then
      CurrentQuery.FieldByName("CARTAREMESSA").Value = VALOR
    End If
    FinalizaSamParametrosProcContas

    If NodeInternalCode = 10 Then
      CurrentQuery.FieldByName("PEG").Clear
      CurrentQuery.FieldByName("CARTAREMESSA").Clear

      PEG.ReadOnly = AtribuirReadOnly(False)
    Else
      PEG.ReadOnly = AtribuirReadOnly(True)
    End If
  Else
    PEG.ReadOnly = AtribuirReadOnly(False)
  End If

  InicializaSamParametrosProcContas("UTILIZAIDENTIFICADORLOTE, MASCARAIDENTIFICADORLOTE")
  If qSamParametrosProcContas.FieldByName("UTILIZAIDENTIFICADORLOTE").AsString = "S" Then
    If Trim(qSamParametrosProcContas.FieldByName("MASCARAIDENTIFICADORLOTE").AsString) <> "" Then
      CurrentQuery.FieldByName("IDENTIFICADORLOTE").Mask = qSamParametrosProcContas.FieldByName("MASCARAIDENTIFICADORLOTE").AsString + ";1;_"
    End If
  End If
  FinalizaSamParametrosProcContas

  If (VisibleMode Or WebMode) Then
    vRegime = VerificaRegimePgto

    If vRegime = 1 Or vRegime = 2 Then
      CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = vRegime
    End If
  End If
End Sub


Public Sub TABLE_AfterPost()
  ' Se é uma inclusão e o Digitador não foi preenchido
  If (Estadodatabela = 3 And CurrentQuery.FieldByName("DIGITADOR").IsNull) Then
	Dim qDigitacao As Object
    Set qDigitacao = NewQuery
	qDigitacao.Clear

    qDigitacao.Add("UPDATE SAM_PEG SET DIGITADOR = :PDIGITADOR, DATADIGITACAO = :PDATA WHERE HANDLE = :PHANDLE")
    qDigitacao.ParamByName("PDIGITADOR").AsInteger = CurrentUser
	qDigitacao.ParamByName("PDATA").AsDateTime = ServerNow
	qDigitacao.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

	qDigitacao.ExecSQL
    Set qDigitacao = Nothing
  End If

  If CurrentQuery.State = 2 Then
    Dim vsMensagem As String
    'Verificar se é permitido alterar o PEG conforme as regras do Provisionamento
    If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem) = 1 Then
      Err.Raise(vbsUserException, "", vsMensagem)
    End If
  End If

  If (CurrentQuery.FieldByName("PRONTOPARAPROVISAO").AsString = "S") Then

    If VisibleMode Then
      Dim spPreparaProvisao As Object
      Set spPreparaProvisao = NewStoredProc

      spPreparaProvisao.AutoMode = True
      spPreparaProvisao.Name = "BS_PREPARAPROVISAOPEG"

      spPreparaProvisao.AddParam("P_HANDLEPEG", ptInput)
      spPreparaProvisao.AddParam("P_USUARIO", ptInput)
      spPreparaProvisao.AddParam("P_PROBLEMA", ptOutput)

      spPreparaProvisao.ParamByName("P_HANDLEPEG").DataType = ftInteger
      spPreparaProvisao.ParamByName("P_USUARIO").DataType = ftInteger
      spPreparaProvisao.ParamByName("P_PROBLEMA").DataType = ftString

      spPreparaProvisao.ParamByName("P_HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      spPreparaProvisao.ParamByName("P_USUARIO").AsInteger = CurrentUser

      spPreparaProvisao.ExecProc

      If Len(spPreparaProvisao.ParamByName("P_PROBLEMA").AsString) > 0 Then
        bsShowMessage(spPreparaProvisao.ParamByName("P_PROBLEMA").AsString, "I")

        Dim qRetornaNaoProntoProvisao As Object
        Set qRetornaNaoProntoProvisao = NewQuery

        qRetornaNaoProntoProvisao.Clear
        qRetornaNaoProntoProvisao.Add("UPDATE SAM_PEG SET PRONTOPARAPROVISAO = 'N' WHERE HANDLE = :HANDLE")
        qRetornaNaoProntoProvisao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qRetornaNaoProntoProvisao.ExecSQL

        Set qRetornaNaoProntoProvisao = Nothing
      End If
      Set spPreparaProvisao = Nothing
    Else
      Dim vsMensagemErro As String
      Dim vvContainer As CSDContainer
      Dim Obj As Object
      Dim viRetorno As Integer

      Set vvContainer = NewContainer

      vvContainer.AddFields("HANDLE:INTEGER;")
      vvContainer.AddFields("USUARIO:INTEGER;")
      vvContainer.Insert
      vvContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      vvContainer.Field("USUARIO").AsInteger = CurrentUser

      Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                                  "SAMPEG", _
	                                  "PrepararParaProvisao", _
	                                  "Prepara o peg para provisionar", _
	                                   0, _
	                                  "SAM_PEG", _
	                                  "", _
	                                  "", _
	                                  "", _
	                                  "", _
	                                  True, _
	                                  vsMensagemErro, _
	                                  vvContainer)

	    If viRetorno = 0 Then
	      bsShowMessage("Processo enviado para execução no servidor!", "I")
	    Else
	      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	    End If

	    Set Obj = Nothing
	End If
  End If


  If (Not CurrentQuery.FieldByName("PEGRASTREADOR").IsNull) Then

      Dim qUpdateRastreamento As Object
      Set qUpdateRastreamento = NewQuery

      qUpdateRastreamento.Add("UPDATE SAM_PEG_RASTREADOR          ")
      qUpdateRastreamento.Add("   SET PEG    = :HANDLEPEG         ")
      qUpdateRastreamento.Add(" WHERE HANDLE = :HANDLERASTREAMENTO")

      qUpdateRastreamento.ParamByName("HANDLEPEG"         ).AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qUpdateRastreamento.ParamByName("HANDLERASTREAMENTO").AsInteger = CurrentQuery.FieldByName("PEGRASTREADOR").AsInteger

      qUpdateRastreamento.ExecSQL

      Set qUpdateRastreamento = Nothing

  End If

  If (EstaNaInterfaceDeDigitacao()) Then
	TABLE_AfterCommitted
  End If

End Sub


Public Sub TABLE_AfterScroll()

  BOTAODIGITAR.Visible       = Not PermitirDigitacaoPeg()
  BOTAOPEGORIGINAL.Visible   = Not CurrentQuery.FieldByName("PEGORIGINAL").IsNull
  BOTAORECLASSIFICAR.Visible = (CurrentQuery.FieldByName("SITUACAO").AsString = "2" Or _
                                CurrentQuery.FieldByName("SITUACAO").AsString = "3" Or _
                                CurrentQuery.FieldByName("SITUACAO").AsString = "6")

  gTabOrigemRecursoPEG = CurrentQuery.FieldByName("TABORIGEMRECURSOPEG").AsInteger

  Set qVerificaConsiderarSp = NewQuery

  qVerificaConsiderarSp.Active = False
  qVerificaConsiderarSp.Clear
  qVerificaConsiderarSp.Add("SELECT CONTROLADOTORC FROM SFN_PARAMETROSFIN")
  qVerificaConsiderarSp.Active = True

  If qVerificaConsiderarSp.FieldByName("CONTROLADOTORC").AsInteger = 1 Then
    TABLE.TabVisible(6) = False
  Else
    TABLE.TabVisible(6) = True
  End If

  Set qVerificaConsiderarSp = Nothing

  SessionVar("hPeg") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger) 'Esse handle é utilizado na Digitação de Guias e na Conferência de Pegs - NÃO DELETAR
  SessionVar("HANDLE_PEG") = CurrentQuery.FieldByName("HANDLE").AsString

  BOTAODIGITAR.Enabled = True

  TIPODEGUIA.WebLocalWhere = "A.TIPOGUIATISS <> 'N'"
  TIPODEGUIA.LocalWhere    = "TIPOGUIATISS <> 'N'"

  If CurrentQuery.State = 1 Then
    TIPODEGUIA.ReadOnly = True
  End If

  vTratarAlteracaoCredito = True
  vPodeSAlvarPegDataPagamentoCorreta = True

  VerificaReciboNF

  If (EhPrestadorFixo(CurrentQuery.FieldByName("RECEBEDOR").AsInteger) And (CurrentQuery.FieldByName("SITUACAO").AsString = "1")) Then
    VALORPAGAMENTORATEIO.ReadOnly = False
  Else
    VALORPAGAMENTORATEIO.ReadOnly = True
  End If

  BOTAOALTERARDOTACAO.Visible = (CurrentQuery.FieldByName("SITUACAO").AsString = "4") And (UtilizaDotacaoOrcamentaria())
  BOTAOALTERARDOTACAO.Visible = BOTAOALTERARDOTACAO.Visible And ((CurrentQuery.FieldByName("TABORIGEMRECURSOPEG").AsInteger = 2) Or (CurrentQuery.FieldByName("TABORIGEMRECURSOPEGCALC").AsInteger = 2))


  If CurrentQuery.FieldByName("SITUACAO").AsString = "6" Then
    DEVOLVERPEG.Enabled =False
    FASEPEG.Enabled = False
    FASEPEGTODOS.Enabled = False
    BOTAOVERALERTAS.Enabled = False
    CONFEVENTO.Enabled = False
    CONFERIDO.Enabled = False
    CONFGUIA.Enabled = False
    CONFPEG.Enabled = False
    DESDOBRAR.Enabled = False
    IMPORTARBENNER.Enabled = False
    IMPORTARSUS.Enabled = False
    REVISAREVENTOS.Enabled = False
    CONTATERCEIRO.Enabled = False
    CRITICARDIGITACAO.Enabled = False
    CONCILIARNOTA.Enabled = False
    BOTAOINCLUIRPRESTADOR.Enabled = False
    BOTAOCANCELAFATURAMENTO.Enabled = False
    REPROCESSARPEG.Enabled = True
    BOTAOLIBERARVERIFICACAO.Enabled = True
  Else
    DEVOLVERPEG.Enabled =True
    FASEPEG.Enabled = True
    FASEPEGTODOS.Enabled = True
    BOTAOVERALERTAS.Enabled = True
    CONFEVENTO.Enabled = True
    CONFERIDO.Enabled = True
    CONFGUIA.Enabled = True
    CONFPEG.Enabled = True
    DESDOBRAR.Enabled = True
    IMPORTARBENNER.Enabled = True
    IMPORTARSUS.Enabled = True
    REVISAREVENTOS.Enabled = True
    CONTATERCEIRO.Enabled = True
    CRITICARDIGITACAO.Enabled = True
    CONCILIARNOTA.Enabled = True
    BOTAOINCLUIRPRESTADOR.Enabled = True
    BOTAOCANCELAFATURAMENTO.Enabled = True
    REPROCESSARPEG.Enabled = True
    BOTAOLIBERARVERIFICACAO.Enabled = False
  End If

  If (NodeInternalCode = 5) Or (EstaNaInterfaceDeDigitacao()) Then
    BOTAOINCLUIRPRESTADOR.Enabled = False
    BOTAOCOMPENTREGUEOUTROPLANO.Enabled = False
  End If

  If (((NodeInternalCode = 2) Or (EstaNaInterfaceDeDigitacao())) And (CurrentQuery.FieldByName("SITUACAO").AsString = "4")) Then
	BOTAOINCLUIROBS.Enabled = True
  Else
	BOTAOINCLUIROBS.Enabled = False
	Dim vsMensagemAux As String
    If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagemAux) = 1 Then
       BOTAOINCLUIROBS.Enabled = True
    End If
  End If

  InicializaSamParametrosProcContas("UTILIZAIDENTIFICADORLOTE, MASCARAIDENTIFICADORLOTE")
  If qSamParametrosProcContas.FieldByName("UTILIZAIDENTIFICADORLOTE").AsString ="S" Then
    If (CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "2" Or CurrentQuery.FieldByName("SITUACAO").AsString = "3") Then
      If NodeInternalCode <> 5 Then
        CurrentQuery.FieldByName("IDENTIFICADORLOTE").Mask = qSamParametrosProcContas.FieldByName("MASCARAIDENTIFICADORLOTE").AsString
        IDENTIFICADORLOTE.ReadOnly = AtribuirReadOnly(False)
      End If
    End If
  Else
  	 IDENTIFICADORLOTE.ReadOnly = True
  End If
  FinalizaSamParametrosProcContas

  If Not CurrentQuery.FieldByName("PEG").IsNull Then
    PEG.ReadOnly = AtribuirReadOnly(True)
  Else
    PEG.ReadOnly = AtribuirReadOnly(False)
  End If

  'Em modo Web os campos de conta terceiro serão digitados na própria visão do Peg
  If VisibleMode Then
    BANCO.ReadOnly = AtribuirReadOnly(True)
    AGENCIA.ReadOnly = AtribuirReadOnly(True)
    CONTACORRENTENUMERO.ReadOnly = AtribuirReadOnly(True)
    CONTACORRENTEDV.ReadOnly = AtribuirReadOnly(True)
    CONTACORRENTENOME.ReadOnly = AtribuirReadOnly(True)
    CONTACORRENTECPFCNPJ.ReadOnly = AtribuirReadOnly(True)
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString < "3" Then
    RECEBEDOR.ReadOnly = AtribuirReadOnly(False)
    LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(False)
  Else
    RECEBEDOR.ReadOnly = AtribuirReadOnly(True)
    LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(True)
  End If

  Dim q1 As Object
  Set q1 = NewQuery

  If (NodeInternalCode = 10) And (CurrentQuery.State = 3) Then
	SEQUENCIA.ReadOnly =AtribuirReadOnly(False)
	PEG.ReadOnly = AtribuirReadOnly(False)
  Else
	PEG.ReadOnly = AtribuirReadOnly(True)
	SEQUENCIA.ReadOnly = AtribuirReadOnly(True)
  End If

  QuatidadeGuia

  If CurrentQuery.FieldByName("TABREGIMEPGTO").Value = 2 Or CurrentQuery.FieldByName("PEGORIGINAL").AsInteger > 0 Then
    GRUPOADIANTAMENTO.Visible =False
  Else
    GRUPOADIANTAMENTO.Visible = True
  End If

  Dim vCOMPETENCIA As Long
  If (VisibleMode Or WebMode) Then
	vCOMPETENCIA = RecordHandleOfTableInterfacePEG("SAM_COMPETPEG")
  Else
	vCOMPETENCIA = CurrentQuery.FieldByName("COMPETENCIA").AsInteger
  End If

  If (vCOMPETENCIA > 0) Then
    FILIAL.ReadOnly = AtribuirReadOnly(True)
    FILIALPROCESSAMENTO.ReadOnly = AtribuirReadOnly(True)
    COMPETENCIA.ReadOnly = AtribuirReadOnly(False)
  End If


  processachange2 = True
  processachange = True
  gchangeValorAdi = True
  gchangeDataAdi = True
  OLDRECEBEDOR = 0

  ARRUMAROTULOS

  q1.Active = False

  q1.Clear

  q1.Add("SELECT SUM(E.VALORPAGTO) VALOR								   ")
  q1.Add("  FROM SAM_PEG P, SAM_GUIA G, SAM_GUIA_EVENTOS E				   ")
  q1.Add(" WHERE P.HANDLE = :PEG AND G.PEG = P.HANDLE AND G.HANDLE = E.GUIA")
  q1.Add("   AND E.COPIAEVENTOORIGINAL <> 'S'                              ")
  q1.ParamByName("PEG").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  q1.Active = True

  ROTULOCALPGTO.Text = "Valor pago calculado [" +Format(q1.FieldByName("VALOR").AsFloat,"###,###,##0.00")+"]"

  q1.Active = False
  Set q1 = Nothing

  If (CurrentQuery.FieldByName("SITUACAO").AsString = "1") Then ' digitação
    REPROCESSARPEG.Visible = False

    If NodeInternalCode <> 2 Then
      REPROCESSARPEG.Enabled = False
    End If

    If CurrentQuery.State <> 3 Then
	  InicializaSamParametrosProcContas("NAOCRITICARIMPORTACAO")
      If (qSamParametrosProcContas.FieldByName("NAOCRITICARIMPORTACAO").AsString = "S") Then
        If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) Then
          RECEBEDOR.ReadOnly = AtribuirReadOnly(False)
          LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(False)
          BENEFICIARIO.ReadOnly = AtribuirReadOnly(True)
        Else
          RECEBEDOR.ReadOnly = AtribuirReadOnly(True)
          LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(True)
          BENEFICIARIO.ReadOnly = AtribuirReadOnly(False)
        End If
      Else
        RECEBEDOR.ReadOnly = AtribuirReadOnly(True)
        LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(True)
        BENEFICIARIO.ReadOnly = AtribuirReadOnly(True)
      End If
	  FinalizaSamParametrosProcContas
    End If

  Else
    If NodeInternalCode <> 2 Then
      REPROCESSARPEG.Enabled = True
    End If

    If CurrentQuery.State <> 3 Then
      RECEBEDOR.ReadOnly = AtribuirReadOnly(True)
      LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(True)
      BENEFICIARIO.ReadOnly = AtribuirReadOnly(True)
    End If
  End If

  If ((CurrentQuery.FieldByName("SITUACAO").AsString = "4") Or  _
	  (CurrentQuery.FieldByName("SITUACAO").AsString = "5")) Then
    DEVOLVERPEG.Enabled = False
    FASEPEG.Enabled = False
    FASEPEGTODOS.Enabled = False
    BOTAOVERALERTAS.Enabled = False
    CONFEVENTO.Enabled = False
    CONFERIDO.Enabled = False
    CONFGUIA.Enabled = False
    CONFPEG.Enabled = False
    DESDOBRAR.Enabled = False
    IMPORTARBENNER.Enabled = False
    IMPORTARSUS.Enabled = False
    REPROCESSARPEG.Enabled = False
    REVISAREVENTOS.Enabled = False
    CONTATERCEIRO.Enabled = False
    CRITICARDIGITACAO.Enabled = False
    CONCILIARNOTA.Enabled = False
    BOTAOINCLUIRPRESTADOR.Enabled = False
  End If

  If (NodeInternalCode <> 611) And (NodeInternalCode <> 2) And (CurrentQuery.FieldByName("SITUACAO").AsString <> "9") Then
	BOTAOCANCELAFATURAMENTO.Visible = False
  End If

  InicializaSamParametrosProcContas("ORIGEMREEMBOLSO")
  If qSamParametrosProcContas.FieldByName("ORIGEMREEMBOLSO").AsString = "2" Then 'Executor
    ESTADO.ReadOnly = AtribuirReadOnly(True)
    MUNICIPIO.ReadOnly = AtribuirReadOnly(True)
  Else
    ESTADO.ReadOnly = AtribuirReadOnly(False)
    MUNICIPIO.ReadOnly = AtribuirReadOnly(False)
  End If
  FinalizaSamParametrosProcContas

  If (NodeInternalCode <> 611) Then
    VOLTARSITUACAO.Visible = False
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "9" Then
    DEVOLVERPEG.Enabled = False
    FASEPEG.Enabled = False
    FASEPEGTODOS.Enabled = False
    BOTAOVERALERTAS.Enabled = False
    CONFEVENTO.Enabled = False
    CONFERIDO.Enabled = False
    CONFGUIA.Enabled = False
    CONFPEG.Enabled = False
    DESDOBRAR.Enabled = False
    IMPORTARBENNER.Enabled = False
    IMPORTARSUS.Enabled = False
    REPROCESSARPEG.Enabled = False
    REVISAREVENTOS.Enabled = False
    CONTATERCEIRO.Enabled = False
    CRITICARDIGITACAO.Enabled = False
    CONCILIARNOTA.Enabled = False
    BOTAOINCLUIRPRESTADOR.Enabled = False
    BOTAOCANCELAFATURAMENTO.Enabled = False
  End If

  Dim vHabilidaBotao As Boolean

  If (CurrentQuery.FieldByName("SITUACAO").AsString ="8") Or (CurrentQuery.FieldByName("SITUACAO").AsString ="4") Then
    vHabilidaBotao = False
  Else
    vHabilidaBotao =True
  End If

  If CurrentQuery.State = 1 Then
    TableReadOnly = Not vHabilidaBotao
  End If

  DEVOLVERPEG.Enabled = vHabilidaBotao
  FASEPEG.Enabled = vHabilidaBotao
  FASEPEGTODOS.Enabled = vHabilidaBotao
  BOTAOVERALERTAS.Enabled = vHabilidaBotao
  CONFEVENTO.Enabled = vHabilidaBotao
  CONFERIDO.Enabled = vHabilidaBotao
  CONFGUIA.Enabled = vHabilidaBotao
  CONFPEG.Enabled = vHabilidaBotao
  DESDOBRAR.Enabled = vHabilidaBotao
  IMPORTARBENNER.Enabled = vHabilidaBotao
  IMPORTARSUS.Enabled = vHabilidaBotao
  REVISAREVENTOS.Enabled = vHabilidaBotao
  CONTATERCEIRO.Enabled = vHabilidaBotao
  CRITICARDIGITACAO.Enabled = vHabilidaBotao
  CONCILIARNOTA.Enabled = vHabilidaBotao
  BOTAOINCLUIRPRESTADOR.Enabled = vHabilidaBotao
  BOTAOCANCELAFATURAMENTO.Enabled = vHabilidaBotao
  BOTAOCANCELAFATURAMENTO.Enabled = vHabilidaBotao

  HabilitaEdicaoeBotoes

  vSituacaoAnteriorPeg = CurrentQuery.FieldByName("SITUACAO").AsString

  If CurrentQuery.FieldByName("SITUACAO").AsString = "3" Then
  	DESDOBRAR.Enabled = False
  	REVISAREVENTOS.Enabled = False
  End If

    Dim vFaturasComProvisao As Integer
  Dim vFaturasSemProvisao As Integer
  Dim qRotuloProvisao As Object
  Set qRotuloProvisao =NewQuery

  vFaturasComProvisao = 0
  vFaturasSemProvisao = 0


  qRotuloProvisao.Active =False
  qRotuloProvisao.Clear

  qRotuloProvisao.Add(" Select GE.FATURAPROVISAO, COUNT(1) QTDE            ")
  qRotuloProvisao.Add("   FROM SAM_PEG P                                   ")
  qRotuloProvisao.Add("   Join SAM_GUIA G On (P.Handle = G.PEG)            ")
  qRotuloProvisao.Add("   Join SAM_GUIA_EVENTOS GE On (GE.GUIA = G.Handle) ")
  qRotuloProvisao.Add("  WHERE GE.FATURAPROVISAO IS NOT Null               ")
  qRotuloProvisao.Add("    AND P.HANDLE = :HANDLE                          ")
  qRotuloProvisao.Add("  GROUP BY GE.FATURAPROVISAO                        ")
  qRotuloProvisao.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger
  qRotuloProvisao.Active =True
  qRotuloProvisao.First

  If Not qRotuloProvisao.FieldByName("QTDE").IsNull Then
	  vFaturasComProvisao = qRotuloProvisao.FieldByName("QTDE").AsInteger
  Else
	  vFaturasComProvisao = 0
  End If

  qRotuloProvisao.Active =False
  qRotuloProvisao.Clear
  qRotuloProvisao.Add(" Select GE.FATURAPROVISAO, COUNT(1) QTDE            ")
  qRotuloProvisao.Add("   FROM SAM_PEG P                                   ")
  qRotuloProvisao.Add("   Join SAM_GUIA G On (P.Handle = G.PEG)            ")
  qRotuloProvisao.Add("   Join SAM_GUIA_EVENTOS GE On (GE.GUIA = G.Handle) ")
  qRotuloProvisao.Add("  WHERE GE.FATURAPROVISAO IS Null                   ")
  qRotuloProvisao.Add("    AND P.HANDLE = :HANDLE                          ")
  qRotuloProvisao.Add("  GROUP BY GE.FATURAPROVISAO                        ")

  qRotuloProvisao.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("HANDLE").AsInteger

  qRotuloProvisao.Active =True
  qRotuloProvisao.First

  If Not qRotuloProvisao.FieldByName("QTDE").IsNull Then
	vFaturasSemProvisao = qRotuloProvisao.FieldByName("QTDE").AsInteger
  Else
    vFaturasSemProvisao = 0
  End If

  'Não exibir no autorizador externo
  If ((CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) And (Not WebVisionCode = "W_SAM_PEG")) Then

	If vFaturasComProvisao > 0 And vFaturasSemProvisao = 0 Then
		ROTULOPROVISAO.Text = "Todos os eventos do PEG foram provisionados!"
	End If

	If vFaturasComProvisao = 0 And vFaturasSemProvisao > 0 Then
		ROTULOPROVISAO.Text = "PEG sem eventos provisionados!"
	End If

	If vFaturasComProvisao = 0 And vFaturasSemProvisao = 0 Then
		ROTULOPROVISAO.Text = "PEG sem eventos provisionados!"
	End If

	If vFaturasComProvisao > 0 And vFaturasSemProvisao > 0 Then
		ROTULOPROVISAO.Text = "O PEG possui " + Str(      vFaturasComProvisao) + " eventos provisionados e " + Str(vFaturasSemProvisao) + " eventos não provisionados!"
	End If
  End If

  Dim vsMensagem As String

  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
  	BOTAOALTERARDATAPAGAMENTO.Enabled = False
  	BOTAOALTERARGUIASAPRESENTADAS.Enabled = False
  	BOTAOALTERARVALORAPRESENTADO.Enabled = False
  Else
  	BOTAOALTERARDATAPAGAMENTO.Enabled = True
  	BOTAOALTERARGUIASAPRESENTADAS.Enabled = True
  	BOTAOALTERARVALORAPRESENTADO.Enabled = True
  End If

  qRotuloProvisao.Active =False
  Set qRotuloProvisao =Nothing

  If ((CurrentQuery.FieldByName("SITUACAO").Value = 1) Or (CurrentQuery.FieldByName("SITUACAO").Value = 2) Or (CurrentQuery.State = 3)) Then
    REGRACALCBASEISS.ReadOnly = AtribuirReadOnly(False)
    REGRACALCBASECONTFEDERAIS.ReadOnly = AtribuirReadOnly(False)
  Else
    REGRACALCBASEISS.ReadOnly = AtribuirReadOnly(True)
    REGRACALCBASECONTFEDERAIS.ReadOnly = AtribuirReadOnly(True)
  End If

  Set qVerificaConsiderarSp = NewQuery

  qVerificaConsiderarSp.Active = False
  qVerificaConsiderarSp.Clear
  qVerificaConsiderarSp.Add("SELECT CONSIDERARCODSERVICO FROM SFN_PARAMETROSFIN")
  qVerificaConsiderarSp.Active = True

  If (qVerificaConsiderarSp.FieldByName("CONSIDERARCODSERVICO").AsString = "N")  Then
    CODIGOSERVICO.Visible = False
    LISTASERVICO.Visible = False
  Else
    CODIGOSERVICO.Visible = True
    LISTASERVICO.Visible = True
  End If

  vTipoPegAnterior = CurrentQuery.FieldByName("TIPOPEG").AsInteger
  vBeneficiarioAnterior = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger

  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 0 Then
    VTabRegimePgtoAnterior = 1
  Else
    VTabRegimePgtoAnterior = CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger
  End If

  Set qVerificaConsiderarSp = Nothing

  InicializaSamParametrosProcContas("PERMITIRIMPOSTOSNOPEG")
  If (qSamParametrosProcContas.FieldByName("PERMITIRIMPOSTOSNOPEG").AsString = "S") Then ' marcado para alterar
    If (CurrentQuery.FieldByName("SITUACAO").AsString <> "4") Then ' não faturado
      BASERETENCAOISS.ReadOnly = AtribuirReadOnly(False)
      BASERETENCAOINSS.ReadOnly = AtribuirReadOnly(False)
      BASERETENCAOIRRF.ReadOnly = AtribuirReadOnly(False)
      BASERETENCAOCONTRIBSOCIAIS.ReadOnly = AtribuirReadOnly(False)
    Else
      BASERETENCAOISS.ReadOnly = True
      BASERETENCAOINSS.ReadOnly = True
      BASERETENCAOIRRF.ReadOnly = True
      BASERETENCAOCONTRIBSOCIAIS.ReadOnly = True
    End If
  Else
    BASERETENCAOISS.ReadOnly = True
    BASERETENCAOINSS.ReadOnly = True
    BASERETENCAOIRRF.ReadOnly = True
    BASERETENCAOCONTRIBSOCIAIS.ReadOnly = True
  End If

  FinalizaSamParametrosProcContas

  If (CurrentQuery.FieldByName("AGRUPADORPAGAMENTO").AsString <> "") Then
	RECEBEDOR.ReadOnly = True
  End If

  HabilitarBotoesTriagem

  BOTAOALTERARDATACONTABIL.Enabled = ( vFaturasComProvisao < 1 ) And (CurrentQuery.FieldByName("SITUACAO").AsString <> "4" )
  BOTAOCANCELARPROVISAO.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString <> "4"
  BOTAOVERIFICAMONITORAMENTO.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString <> "4"
  BOTAOPROVISIONARPEG.Enabled =  CurrentQuery.FieldByName("SITUACAO").AsString <> "4"
  BOTAOALTERARIDENTIFICADORPAGTO.Enabled =  CurrentQuery.FieldByName("SITUACAO").AsString <> "4"


  Dim RotinaAdiantamento As Object

  Set RotinaAdiantamento = NewQuery

  RotinaAdiantamento.Clear
  RotinaAdiantamento.Active = False
  RotinaAdiantamento.Add(" SELECT HANDLE    	    			    ")
  RotinaAdiantamento.Add("   FROM SFN_ROTINAFINADIANT 				")
  RotinaAdiantamento.Add("  WHERE :PEG BETWEEN PEGINICIAL And PEGFINAL")
  RotinaAdiantamento.Add("    AND SITUACAO = 5						")
  RotinaAdiantamento.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  RotinaAdiantamento.Active = True

  If Not (RotinaAdiantamento.EOF) Then
    VALORADIANTAMENTO.ReadOnly = True
    VALORDESCONTO.ReadOnly = True
    DATAADIANTAMENTO.ReadOnly = True
    ADIANTAMENTO.ReadOnly = True
  End If

  Set RotinaAdiantamento = Nothing



  If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsString = "2") Then

    If (CurrentQuery.FieldByName("SITUACAO").AsString = "1") Or (CurrentQuery.FieldByName("SITUACAO").AsString = "2") Then

      SetarReadOnlyRecursoProprioOrcamento(False)

    Else

      SetarReadOnlyRecursoProprioOrcamento(True)

    End If

  End If

  If Not CurrentQuery.FieldByName("HANDLE").IsNull Then
    If VerificarBloqueioPegExisteComposicao(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
      BOTAOALTERARIDENTIFICADORPAGTO.Enabled = False
    Else
      BOTAOALTERARIDENTIFICADORPAGTO.Enabled = True
    End If

    If VerificarPegReapresentadoExercicioPosterior(CurrentQuery.FieldByName("HANDLE").AsInteger) And CurrentQuery.FieldByName("SITUACAO").AsString = "1" Then
      BOTAOALTERAREMPENHO.Enabled = True
    Else
      BOTAOALTERAREMPENHO.Enabled = False
    End If

	If VerificarBloqueioAlteracoesReapresentacao(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
	  BOTAODIGITAR.Enabled = False
      DESDOBRAR.Enabled = False
	  IMPORTARBENNER.Enabled = False
	  IMPORTARSUS.Enabled = False
	  CONTATERCEIRO.Enabled = False
	  BOTAOALTERARGUIASAPRESENTADAS.Enabled = False
	  BOTAOALTERARVALORAPRESENTADO.Enabled = False
	  BOTAOCANCELARPROVISAO.Enabled = False
	  BOTAOPROVISIONARPEG.Enabled = False
	  BOTAOINCLUIROBS.Enabled = False
	  CONCILIARNOTA.Enabled = False
	End If

	If VerificarBloqueioAlteracoes(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
	  BOTAODIGITAR.Enabled = False
      DESDOBRAR.Enabled = False
	  IMPORTARBENNER.Enabled = False
	  IMPORTARSUS.Enabled = False
	  CONTATERCEIRO.Enabled = False
	  BOTAOALTERARDATAPAGAMENTO.Enabled = False
	  BOTAOALTERARGUIASAPRESENTADAS.Enabled = False
	  BOTAOALTERARVALORAPRESENTADO.Enabled = False
	  BOTAOCANCELARPROVISAO.Enabled = False
	  BOTAOPROVISIONARPEG.Enabled = False
	End If
  End If

End Sub

Public Sub SetarReadOnlyRecursoProprioOrcamento(somenteLeitura As Boolean)
      TABORIGEMPGTOPEG.ReadOnly = somenteLeitura
      TABORIGEMRECURSOPEG.ReadOnly = somenteLeitura
      EMPENHOPEG.ReadOnly = somenteLeitura
      DOTACAOEXERCICIOPEG.ReadOnly = somenteLeitura
      DOTACAONATUREZAPEG.ReadOnly = somenteLeitura
      DOTACAOPEG.ReadOnly = somenteLeitura
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  bsShowMessage("Somente possível excluir PEG através do botão Excluir PEG", "E")
  CanContinue = False
  Exit Sub
End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If VerificarBloqueioAlteracoesReapresentacao(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("O PEG não pode ser alterado porque é de reapresentação. ", "E")
    CanContinue = False
	Exit Sub
  End If

  If VerificarBloqueioAlteracoes(CurrentQuery.FieldByName("HANDLE").AsInteger) Then
    bsShowMessage("O PEG não pode ser alterado porque está vinculado a um agrupador de pagamento com documentos fiscais conciliados. ", "E")
    CanContinue = False
	Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    bsShowMessage("PEG já pago não  pode ser alterado", "I")
    CanContinue =False
  End If

	If Not(VerificarUsuarioParaEdicao) Then
		CanContinue = False
		Exit Sub
	End If
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  Dim vsMensagem As String

  'Verificar se é permitido alterar o PEG conforme as regras do Provisionamento
  If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem) = 1 Then
    CanContinue = False
    bsShowMessage(vsMensagem, "E")
    Exit Sub
  End If

  CanContinue = CheckFilialProcessamento(CurrentSystem, CurrentQuery.FieldByName("FILIAL").AsInteger, "A")

  If CanContinue = False Then
    AtualizarCarga(True)
    Exit Sub
  End If
  OLDDATARECEBIMENTO = CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime
  OLDDATAPAGAMENTO = CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime
  OLDRECEBEDOR = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  OLDLOCALEXECUCAO = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
  vOLDBeneficiarioTitular = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  OldQtdGuia = CurrentQuery.FieldByName("QTDGUIA").AsInteger 'LARINI CAMED
  vbEditando = True
  viHTipoPegAnterior = CurrentQuery.FieldByName("TIPOPEG").AsInteger
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
	vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
	vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If (vFilial > 0) Then
    CanContinue = CheckFilialProcessamento(CurrentSystem, vFilial, "I")

    If CanContinue = False Then
			AtualizarCarga(True)
      Exit Sub
    End If

    vgFilial = vFilial
    vgFilialProcessamento = BuscarFilialProcessamento(CurrentSystem, vFilial)
  Else
    'Segundo o Celso Lara, deve-se na pasta de "Digitação de pegs" considerar a filial padrão do usuário
    'Desta forma, no caso de pegs de reembolso, a filial padrão será a do usuário mesmo
    'e no caso de peg de credenciamento será a filial padrão do recebedor.
    'Rodrigo Postai - 3/08/2005
    Dim sql As Object
    Set sql = NewQuery

    sql.Clear

    sql.Add("SELECT FILIALPADRAO FROM Z_GRUPOUSUARIOS WHERE HANDLE = :HANDLE")

    sql.ParamByName("HANDLE").AsInteger = CurrentUser
    sql.Active = True

    If Not sql.FieldByName("FILIALPADRAO").IsNull Then
      vgFilial = sql.FieldByName("FILIALPADRAO").AsInteger
      vgFilialProcessamento = BuscarFilialProcessamento(CurrentSystem,sql.FieldByName("FILIALPADRAO").AsInteger)
      CanContinue = CheckFilialProcessamento(CurrentSystem, vgFilialProcessamento, "I")

      If CanContinue = False Then
				AtualizarCarga(True)
      	Exit Sub
      End If
    Else
      bsShowMessage("Filial de processamento Nula. É necessário que exista uma filial padrão definida no cadastro do usuário.", "I")
      vgFilial = 0
      vgFilialProcessamento = 0

      FILIAL.ReadOnly = AtribuirReadOnly(False) 'sms 63627 - Edilson.Castro - 14/06/2006
      FILIALPROCESSAMENTO.ReadOnly = AtribuirReadOnly(False) 'sms 63627 - Edilson.Castro - 14/06/2006

      CanContinue = False
    End If

    Set sql = Nothing
  End If

  Dim vCOMPETENCIA As Long
  If (VisibleMode Or WebMode) Then
	vCOMPETENCIA = RecordHandleOfTableInterfacePEG("SAM_COMPETPEG")
  Else
	vCOMPETENCIA = CurrentQuery.FieldByName("COMPETENCIA").AsInteger
  End If

  If ( vCOMPETENCIA > 0) Then
    Dim q1 As Object
    Set q1 = NewQuery

    q1.Add("SELECT SITUACAO FROM SAM_COMPETPEG WHERE HANDLE=:HANDLE")

    q1.ParamByName("HANDLE").Value = vCOMPETENCIA
    q1.Active = True

    If q1.FieldByName("SITUACAO").AsString = "F" Then
      bsShowMessage("Não é possível incluir PEG em competência fechada!", "I")
      CanContinue = False
    End If
  Else
    COMPETENCIA.ReadOnly = AtribuirReadOnly(False) 'sms 63627 - Edilson.Castro - 14/06/2006
  End If
End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If Not CurrentQuery.FieldByName("NFNUMERO").IsNull Then
    CurrentQuery.FieldByName("NFNUMERO").AsString = LimpaEspaco(CurrentQuery.FieldByName("NFNUMERO").AsString)
  End If

  If Not ValidarEmpenhoConformeRecebedor() Then
    bsShowMessage("Informado Empenho Orçamentário incompatível com Recebedor", "E")
    CanContinue = False
    Exit Sub
  End If

  Estadodatabela = CurrentQuery.State

  If (VerificarRecebedorUtilizaAdiantamentoAutomatico) Then
    PreencherAdiantamentoDigitacaoPeg
  End If

  Dim agrupadorFechado As Boolean
  agrupadorFechado = VerificaAgrupadorPagamentoFechado

  If (agrupadorFechado) Then
    BsShowMessage("Não é permitida a alteração do PEG que está ligado à registro de pagamento fechado.","E")
    CanContinue = False
    Exit Sub
  End If

  If TABLE.TabVisible(6) Then
    If CurrentQuery.FieldByName("TABORIGEMPGTOPEG").AsInteger = 1 Then
      CurrentQuery.FieldByName("DOTACAOEXERCICIOPEGCALC").Clear
      CurrentQuery.FieldByName("DOTACAOPEGCALC").Clear
      CurrentQuery.FieldByName("DOTACAONATUREZAPEGCALC").Clear
      CurrentQuery.FieldByName("EMPENHOPEGCALC").Clear
      CurrentQuery.FieldByName("TABORIGEMRECURSOPEGCALC").Clear

      If CurrentQuery.FieldByName("TABORIGEMRECURSOPEG").AsInteger = 1 Then
        CurrentQuery.FieldByName("DOTACAOEXERCICIOPEG").Clear
        CurrentQuery.FieldByName("DOTACAOPEG").Clear
        CurrentQuery.FieldByName("DOTACAONATUREZAPEG").Clear
        CurrentQuery.FieldByName("EMPENHOPEG").Clear
      End If
    ElseIf CurrentQuery.FieldByName("TABORIGEMPGTOPEG").AsInteger = 2 Then
      CurrentQuery.FieldByName("DOTACAOEXERCICIOPEG").Clear
      CurrentQuery.FieldByName("DOTACAOPEG").Clear
      CurrentQuery.FieldByName("DOTACAONATUREZAPEG").Clear
      CurrentQuery.FieldByName("EMPENHOPEG").Clear
      CurrentQuery.FieldByName("TABORIGEMRECURSOPEG").Clear
    End If
  End If

  If (CurrentQuery.FieldByName("DATAADIANTAMENTO").IsNull) And (CurrentQuery.FieldByName("ADIANTAMENTO").AsBoolean)Then
    bsShowMessage("Deve ser informada uma data de adiantamento", "E")
    CanContinue = False
    Exit Sub
  End If

  InicializaSamParametrosProcContas("OBRIGATITULARREEMBOLSO")
  If qSamParametrosProcContas.FieldByName("OBRIGATITULARREEMBOLSO").AsString = "S" Then
    If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2) Then
      Dim Especifico As Object
      Set Especifico = CreateBennerObject("Especifico.uEspecifico")
      If (Especifico.Cliente(CurrentSystem) <> cCabesp) And (CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = 0) Then
        bsShowMessage("O campo Beneficiário Titular é obrigatório para operações de reembolso!", "E")
        CanContinue = False
        Exit Sub
      End If
      Set Especifico = Nothing

      If ((CurrentQuery.FieldByName("BENEFICIARIO").AsInteger) <> vOLDBeneficiarioTitular) And (vOLDBeneficiarioTitular > 0) Then
        If (WebMode) Then
          ApagarBeneficiariosGuiaEventos
	    Else
          If bsShowMessage("Ao alterar o Beneficiário Titular do PEG todos os beneficiários da(s) guia(s) e do(s) evento(s) serão excluídos, deseja continuar?", "Q") = vbNo Then
            CanContinue = False
            Exit Sub
          Else
            ApagarBeneficiariosGuiaEventos
          End If
        End If
      End If
    End If
  End If
  FinalizaSamParametrosProcContas

  If Not ValidaNumeroCartaRemessa Then
    CanContinue = False
    Exit Sub
  End If

  Dim qRegraBaseISS As Object
  Dim qRegraBaseContribFederais As BPesquisa

  If CurrentQuery.FieldByName("DATACONTABIL").IsNull Then
    CurrentQuery.FieldByName("DATACONTABIL").AsDateTime = CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime
  End If

  If CurrentQuery.State = 2 Then
    Dim vsMensagem As String

    'Verificar se é permitido alterar o PEG conforme as regras do Provisionamento
    If PermissaoAlteracao(CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0, True, vsMensagem) = 1 Then
      CanContinue = False
      bsShowMessage(vsMensagem, "E")
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
	  If CurrentQuery.FieldByName("LOCALEXECUCAO").IsNull Then
	  	bsShowMessage("O campo local de execução é obrigatório", "E")
	  	CanContinue = False
	  	Exit Sub
	  End If
  End If

  If CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
    If Not VERIFICAPAG("E") Then
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
  	bsShowMessage("O campo data pagamento é obrigatório", "E")
  	CanContinue = False
  	Exit Sub
  End If

  If CurrentQuery.FieldByName("NUMEROPAGAMENTO").IsNull Then
  	bsShowMessage("O campo número pagamento é obrigatório", "E")
  	CanContinue = False
  	Exit Sub
  End If

  Dim vCOMPETENCIA As Long
  If (VisibleMode Or WebMode) Then
    vCOMPETENCIA = RecordHandleOfTableInterfacePEG("SAM_COMPETPEG")
  Else
    vCOMPETENCIA = CurrentQuery.FieldByName("COMPETENCIA").AsInteger
  End If

  If (vCOMPETENCIA <= 0) And CurrentQuery.State <> 1 Then
	Dim sqlx As Object
	Set sqlx = NewQuery

  	sqlx.Add("SELECT HANDLE,COMPETENCIA FROM SAM_COMPETPEG WHERE COMPETENCIA <= :DATA AND SITUACAO = 'A' ORDER BY COMPETENCIA DESC")

	sqlx.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime
	sqlx.Active = True

	If sqlx.FieldByName("HANDLE").AsInteger = 0 Then
	  bsShowMessage("Favor verificar as competências do processamento de contas, pois a data de recebimento" + Chr(13) + _
	  	  "está em uma competência inexistente ou já finalizada", "E")
	  CanContinue = False

	  Set sqlx = Nothing
	  Exit Sub
	Else
	  CurrentQuery.FieldByName("COMPETENCIA").AsInteger = sqlx.FieldByName("HANDLE").AsInteger
	End If

	Set sqlx = Nothing
  End If

  Dim vAlgumPreenchido As Boolean
  Dim Msg As String

  VerificaReciboNF

  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2 And CurrentQuery.FieldByName("BENEFICIARIO").AsInteger > 0 Then
	  Dim sqlfp, sql2, sql3, sql4 As Object
	  Set sql2=NewQuery
	  Set sql3=NewQuery
	  Set sql4=NewQuery

	  sql2.Add("SELECT B.FAMILIA, B.EHTITULAR, B.FILIALCUSTO, ")
	  sql2.Add("       F.TABRESPONSAVEL, F.TITULARRESPONSAVEL ")
	  sql2.Add("  FROM SAM_BENEFICIARIO B                     ")
	  sql2.Add("  JOIN SAM_FAMILIA F ON F.HANDLE = B.FAMILIA  ")
	  sql2.Add(" WHERE B.HANDLE=:HANDLE                       ")
	  sql2.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
	  sql2.Active=True

      If sql2.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then ' o responsável é beneficiario
        If sql2.FieldByName("TITULARRESPONSAVEL").AsInteger > 0 Then ' está preenchido o campo beneficiário na família
          sql3.Clear
          sql3.Add("SELECT EHTITULAR FROM SAM_BENEFICIARIO WHERE HANDLE=:HANDLE")
          sql3.ParamByName("HANDLE").AsInteger = sql2.FieldByName("TITULARRESPONSAVEL").AsInteger
          sql3.Active = True

          If sql3.FieldByName("EHTITULAR").AsString <> "S" Then ' o beneficiário encontrato não é titular, então faz como fazia antes, usa o titular da família do próprio benef
            sql4.Clear
            sql4.Add("SELECT HANDLE FROM SAM_BENEFICIARIO WHERE FAMILIA=:FAMILIA AND EHTITULAR='S'")
            sql4.ParamByName("FAMILIA").AsInteger = sql2.FieldByName("FAMILIA").AsInteger
            sql4.Active = True
            If sql4.FieldByName("HANDLE").AsInteger > 0 Then 'achou um beneficiário titular na própria família, então usa ele, se não achou, não faz nada
              CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = sql4.FieldByName("HANDLE").AsInteger
            End If
  	      Else ' o beneficiário encontrato é titular, então atualiza ele no peg
            CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = sql2.FieldByName("TITULARRESPONSAVEL").AsInteger
  	      End If
  	    Else 'não está preenchido o campo beneficiario na família então pega o handle do titular responsável da família do próprio beneficiário
          sql4.Clear
          sql4.Add("SELECT HANDLE FROM SAM_BENEFICIARIO WHERE FAMILIA=:FAMILIA AND EHTITULAR='S'")
          sql4.ParamByName("FAMILIA").AsInteger = sql2.FieldByName("FAMILIA").AsInteger
          sql4.Active = True
          If sql4.FieldByName("HANDLE").AsInteger > 0 Then 'achou um beneficiário titular na própria família, então usa ele, se não achou, não faz nada
            CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = sql4.FieldByName("HANDLE").AsInteger
          End If
	    End If
	  End If ' se for pessoa nem precisa fazer nada
	  Set sql2=Nothing
	  Set sql3=Nothing
	  Set sql4=Nothing

    If (vgFilial = 0) And (CurrentQuery.FieldByName("FILIAL").IsNull) Then
      bsShowMessage("Filial de Custo do beneficiário nula. É necessário que exista uma filial de custo" + Chr(13) + _
      	  "definida no cadastro de beneficiário.", "E")
      CurrentQuery.FieldByName("FILIAL").Value = Null
	  CanContinue = False
	  Exit Sub
    Else
      If (vgFilialProcessamento = 0) And (CurrentQuery.FieldByName("FILIALPROCESSAMENTO").IsNull) Then
        bsShowMessage("Filial de Custo do beneficiário sem filial de processamento", "E")
        CurrentQuery.FieldByName("FILIALPROCESSAMENTO").Value = Null
        CanContinue = False
        Exit Sub
      End If
    End If
  End If

  'Faz a verificacao se todos os campos estão preenchidos, pois ao preencher um, todos os campos da conta corrente do peg devem ser preenchidos.
  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2 Then
    If ( (CurrentQuery.FieldByName("CREDITOCONTATERCEIROS").AsString = "S") And WebMode) Or _
       (Not(CurrentQuery.FieldByName("BANCO").IsNull)) Or _
       (Not(CurrentQuery.FieldByName("AGENCIA").IsNull)) Or _
       (Not(CurrentQuery.FieldByName("CONTACORRENTENUMERO").IsNull)) Or _
       (Not(CurrentQuery.FieldByName("CONTACORRENTEDV").IsNull)) Or _
       (Not(CurrentQuery.FieldByName("CONTACORRENTENOME").IsNull)) Or _
       (Not(CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").IsNull)) Then
      vAlgumPreenchido = True
    End If

    If ((CurrentQuery.FieldByName("BANCO").IsNull) Or _
        (CurrentQuery.FieldByName("AGENCIA").IsNull) Or _
        (CurrentQuery.FieldByName("CONTACORRENTENUMERO").IsNull) Or _
        (CurrentQuery.FieldByName("CONTACORRENTEDV").IsNull) Or _
        (CurrentQuery.FieldByName("CONTACORRENTENOME").IsNull) Or _
        (CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").IsNull) ) And _
       (vAlgumPreenchido) Then
      bsShowMessage("Todos os campos da conta de terceiro devem ser preenchidos!", "E")

      If CurrentQuery.FieldByName("BANCO").IsNull Then
        BANCO.SetFocus
      ElseIf CurrentQuery.FieldByName("AGENCIA").IsNull Then
        AGENCIA.SetFocus
      ElseIf CurrentQuery.FieldByName("CONTACORRENTENUMERO").IsNull Then
        CONTACORRENTENUMERO.SetFocus
      ElseIf CurrentQuery.FieldByName("CONTACORRENTEDV").IsNull Then
        CONTACORRENTEDV.SetFocus
      ElseIf CurrentQuery.FieldByName("CONTACORRENTENOME").IsNull Then
        CONTACORRENTENOME.SetFocus
      ElseIf CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").IsNull Then
        CONTACORRENTECPFCNPJ.SetFocus
      End If

      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = 0 Then
    CurrentQuery.FieldByName("BENEFICIARIO").Value = Null
  End If

  InicializaSamParametrosProcContas("REPETIRIDENTIFICADORPGTO")

  If qSamParametrosProcContas.FieldByName("REPETIRIDENTIFICADORPGTO").AsString = "N" Then
    Dim vbResultado   As Boolean
    Dim vsMsgVerifica As String
    Dim DLLEspecifico As Object
    Set DLLEspecifico = CreateBennerObject("ESPECIFICO.UESPECIFICO")
    vbResultado = DLLEspecifico.PRO_VerificaIdentificadorPagamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("RECEBEDOR").AsInteger, CurrentQuery.FieldByName("IDENTIFICADORPAGAMENTO").AsString, vsMsgVerifica)
    Set DLLEspecifico = Nothing

    If (vbResultado) Then
      bsshowmessage(vsMsgVerifica, "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim qVerificaTipodePEG As Object
  Set qVerificaTipodePEG = NewQuery

  'Verifica se o regime de pagamento do PEG é compatível com o regime de pagamento do tipo do PEG
  qVerificaTipodePEG.Clear

  qVerificaTipodePEG.Add("SELECT TABREGIMEPGTO,    ")
  qVerificaTipodePEG.Add("       REEMBOLSODECASAL  ")
  qVerificaTipodePEG.Add("  FROM SAM_TIPOPEG       ")
  qVerificaTipodePEG.Add(" WHERE HANDLE = :TIPOPEG ")
  qVerificaTipodePEG.Add("   AND TABREGIMEPGTO <> 3")

  qVerificaTipodePEG.ParamByName("TIPOPEG").AsInteger = CurrentQuery.FieldByName("TIPOPEG").AsInteger
  qVerificaTipodePEG.Active = True

  If (Not qVerificaTipodePEG.EOF) And _
     (qVerificaTipodePEG.FieldByName("TABREGIMEPGTO").AsInteger <> CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger) Then
    bsShowMessage("Regime de pagamento do PEG incompatível com o regime de pagamento do tipo de PEG", "E")
    CanContinue = False
    Exit Sub
  End If

  'Se regime de pagamento é reembolso de casal
  If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2) And (qVerificaTipodePEG.FieldByName("REEMBOLSODECASAL").AsString = "S") Then
    If CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
      bsShowMessage("Para tipo de PEG de reembolso de casal é necessário informar o beneficiário", "E")
      CanContinue = False
      Exit Sub
    End If

    'Pega a informação de relação matrimonial do beneficiário titular
    'Somente permite reembolso de casal se titular tiver cônjuge
	Set DLLEspecifico = CreateBennerObject("ESPECIFICO.UESPECIFICO")

    If Not DLLEspecifico.BEN_TemConjuge(CurrentSystem, CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger) Then
      bsShowMessage("Tipo de PEG de reembolso de casal, mas titular da família do beneficiário não possui relação matrimonial", "E")
      CanContinue = False
  	  Set DLLEspecifico = Nothing
      Exit Sub
    End If
	Set DLLEspecifico = Nothing
  End If

  Set qVerificaTipodePEG = Nothing

  If CurrentQuery.State = 3 Then 'inserir
    If CurrentQuery.FieldByName("FILIALPROCESSAMENTO").IsNull Then
      CurrentQuery.FieldByName("FILIALPROCESSAMENTO").AsInteger = BuscarFilialProcessamento(CurrentSystem, CurrentQuery.FieldByName("FILIAL").AsInteger)
    End If

    Dim QX As Object
    Set QX = NewQuery
    QX.Clear
    InicializaSamParametrosProcContas("REPETICAONUMERACAOPEG")
    Select Case qSamParametrosProcContas.FieldByName("REPETICAONUMERACAOPEG").AsString
	  Case "2"
	    QX.Add("SELECT HANDLE FROM SAM_PEG WHERE PEG=:PEG AND SITUACAO <>'9' AND SEQUENCIA = :SEG")

		QX.ParamByName("SEG").Value = CurrentQuery.FieldByName("SEQUENCIA").AsFloat
		QX.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsFloat
		QX.Active = True

		If Not QX.EOF Then
		  bsShowMessage("Já existe um PEG com este número", "E")
		  PEG.SetFocus
		  CanContinue = False
		  Exit Sub
		End If
 	  Case "3"
		QX.Add("SELECT HANDLE FROM SAM_PEG WHERE COMPETENCIA=:COMPETENCIA AND PEG=:PEG AND SITUACAO <>'9' AND SEQUENCIA = :SEG")

		QX.ParamByName("COMPETENCIA").Value = CurrentQuery.FieldByName("COMPETENCIA").AsFloat
		QX.ParamByName("SEG").Value = CurrentQuery.FieldByName("SEQUENCIA").AsFloat
		QX.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsFloat
		QX.Active = True

		If Not QX.EOF Then
		  bsShowMessage("Já existe um PEG com este número nesta competência", "E")
		  PEG.SetFocus
		  CanContinue = False
		  Exit Sub
		End If
 	  Case "4"
		QX.Add("SELECT HANDLE FROM SAM_PEG WHERE SEQUENCIA = :SEG AND PEG=:PEG AND SITUACAO <>'9' AND FILIAL=:FILIAL")

		QX.ParamByName("SEG").Value = CurrentQuery.FieldByName("SEQUENCIA").AsFloat
		QX.ParamByName("FILIAL").Value = CurrentQuery.FieldByName("FILIAL").AsFloat
		QX.ParamByName("PEG").Value = CurrentQuery.FieldByName("PEG").AsFloat
		QX.Active = True

		If Not QX.EOF Then
		  bsShowMessage("Já existe um PEG com este número nesta filial", "E")
		  PEG.SetFocus
		  CanContinue = False
		  Exit Sub
		End If
    End Select
    FinalizaSamParametrosProcContas

    QX.Active = False
	Set QX = Nothing

    If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
      Dim q3 As Object
      Set q3 = NewQuery
      q3.Clear
      q3.Add("SELECT FILIALPADRAO, DATADESCREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = " +Str(CurrentQuery.FieldByName("RECEBEDOR").AsInteger))
      q3.Active = True

      InicializaSamParametrosProcContas("VERIFICAFILIALPRESTADOR")
      If (qSamParametrosProcContas.FieldByName("VERIFICAFILIALPRESTADOR").AsString = "S")Then
        'Verifica se a filial padrão do prestador é diferente da filial padrão do usuário,caso seja,emitirá uma mensagem
        If (q3.FieldByName("FILIALPADRAO").AsInteger <> CurrentQuery.FieldByName("FILIAL").AsInteger) And _
           (q3.FieldByName("FILIALPADRAO").AsInteger > 0) Then 'sms 28867
          If Not WebMode Then
          	If bsShowMessage("Filial padrão do recebedor é diferente da filial padrão corrente, deseja continuar?", "Q") = vbNo Then
          	    bsShowMessage("Operação Cancelada.", "E")
            	CanContinue = False
            	Exit Sub
          	End If
          End If
        End If
      End If
      FinalizaSamParametrosProcContas

      Set q3 = Nothing
    End If

	If VisibleMode Then
	  If NodeInternalCode = 10 Then
		QX.Active = False

		QX.Clear

		QX.Add("SELECT MIN(SITUACAO) SITUACAO FROM SAM_PEG WHERE PEG=:peg")

		QX.ParamByName("peg").AsInteger = CurrentQuery.FieldByName("PEG").AsFloat
		QX.Active =True

		If (QX.FieldByName("SITUACAO").AsString = "") Then
		  bsShowMessage("Este PEG não é válido.", "E")
		  CanContinue = False
		  PEG.SetFocus
		  Exit Sub
		End If

		If (QX.FieldByName("SITUACAO").AsString <> "4") Then
		  bsShowMessage("Este PEG ou alguma de suas seqüências não se encontra faturado." + Chr(13) + _
		  	  "No Movimento de Acerto de PEGs só é possível incluir sequências de PEGs já faturados", "E")
		  PEG.SetFocus
		  CanContinue = False
		  Exit Sub
		End If

		QX.Active = False

		Set QX = Nothing
	  End If
	End If
  End If

  CanContinue = CheckFilialProcessamento(CurrentSystem,CurrentQuery.FieldByName("FILIAL").AsInteger,"A")

  If CanContinue = False Then
  	bsShowMessage("Problemas na permissão do usuário", "E")
  	AtualizarCarga(False)
    Exit Sub
  End If

  'larini camed
  If CurrentQuery.FieldByName("SITUACAO").AsString <> "1" Then
	If OldQtdGuia <> CurrentQuery.FieldByName("QTDGUIA").AsInteger Then
	  CurrentQuery.FieldByName("SITUACAO").AsString = "2"
	End If
  End If

  If Not DataAdiantamentoOk Then
  	bsShowMessage("A data de adiantamento não está ok", "E")
	CanContinue = False
	Exit Sub
  End If

  Dim q1 As Object
  Dim nguia As Long
  Set q1 = NewQuery

  If Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
	q1.Add("SELECT RECEBEDOR FROM SAM_PRESTADOR WHERE HANDLE=:HANDLE")

	q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
	q1.Active = True

	If q1.FieldByName("RECEBEDOR").AsString <> "S" Then
	  bsShowMessage("O prestador não é recebedor!", "E")
	  CanContinue = False
	  Exit Sub
	End If
  End If

  q1.Clear

  q1.Add("SELECT COUNT(*) NGUIA FROM SAM_GUIA WHERE PEG=:PEG")

  q1.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  q1.Active = True

  nguia = q1.FieldByName("NGUIA").AsInteger

  q1.Active = False

  Set q1 = Nothing

  If (CurrentQuery.FieldByName("QTDGUIA").AsInteger = 0) Or (nguia > CurrentQuery.FieldByName("QTDGUIA").AsInteger) Then
	bsShowMessage("O N. de guias informado deve ser maior que zero e nunca menor que a quantidade já digitada de guias:" + Str(nguia), "E")
	CanContinue = False
	Exit Sub
  End If


  CurrentQuery.FieldByName("DATA").Value = ServerNow
  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser

  OLDDATAPAGAMENTO = -1

  DATARECEBIMENTO_OnExit

  DATAPAGAMENTO_OnExit

  If vVerificaDataWeb Then
  	CanContinue = False
	Exit Sub
  End If

  If Not vPodeSAlvarPegDataPagamentoCorreta Then
  	bsShowMessage("A gravação do peg está impossibilitada pela data de pagamento", "E")
    CanContinue = False
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString <>"4" Then '<> pago
	If CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime <= ServerDate Then
	  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsString ="2" Then 'REEMBOLSO
        If CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
		  bsShowMessage("Informar a data de pagamento no PEG de reembolso", "E")
		  DATAPAGAMENTO.SetFocus
		  CanContinue = False
		  Exit Sub
        Else 'SE A DATA DE PAGAMENTO FOI INFORMADA
		  If CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < ServerDate Then
			bsShowMessage("Entre com data pagamento maior ou igual a de hoje", "E")
            DATAPAGAMENTO.SetFocus
            CanContinue = False
            Exit Sub
		  End If
		End If
	  End If
	Else
	  bsShowMessage("Entre com data de recebimento menor ou igual a de hoje", "E")
	  TABLE.ActivePage(0)
	  DATARECEBIMENTO.SetFocus
	  CanContinue = False
	  Exit Sub
	End If
  End If

  'Verificação de Carta Remessa
  InicializaSamParametrosProcContas("REPETICAONUMERACAOPEG, SUGEREPEGCARTAREMESSA")
  If qSamParametrosProcContas.FieldByName("REPETICAONUMERACAOPEG").Value <> "1" _
     And qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean = True Then

	  If Not CurrentQuery.FieldByName("CARTAREMESSA").IsNull Then
		If CurrentQuery.FieldByName("SEQUENCIA").AsInteger = 0 Then
		  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1 Then
			Dim CARTA2 As Object
			Set CARTA2 = NewQuery

			CARTA2.Active = False
			CARTA2.Clear
			CARTA2.Add("SELECT HANDLE, COMPETENCIA, FILIAL     ")
			CARTA2.Add("  FROM SAM_PEG                         ")
			CARTA2.Add(" WHERE SEQUENCIA = 0                   ")
			CARTA2.Add("   And CARTAREMESSA =  :CARTAREMESSA   ")
			CARTA2.Add("   And FILIAL       =  :FILIAL       and SITUACAO<>'9'   ")
			If CurrentQuery.State <> 3 Then
			  CARTA2.Add("   And HANDLE       <> :HANDLE         ")
			End If
			CARTA2.Add("   And RECEBEDOR    =  :HANDLEPRESTADOR")
			CARTA2.ParamByName("CARTAREMESSA").AsFloat = CurrentQuery.FieldByName("CARTAREMESSA").AsFloat
			CARTA2.ParamByName("FILIAL").Value = CurrentQuery.FieldByName("FILIAL").AsFloat

			If CurrentQuery.State <> 3 Then
			  CARTA2.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
			End If

			CARTA2.ParamByName("HANDLEPRESTADOR").Value = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
			CARTA2.Active = True

			If Not CARTA2.EOF Then
              'Repetir Numeração de Peg em outra Competência
			  If qSamParametrosProcContas.FieldByName("REPETICAONUMERACAOPEG").Value = 3 Then
			    If CARTA2.FieldByName("COMPETENCIA").Value = CurrentQuery.FieldByName("COMPETENCIA").Value Then
				  bsShowMessage("Já existe um PEG com o mesmo número de Carta Remessa desse mesmo Recebedor nesta Competência!", "E")
                  CanContinue = False
                  Exit Sub
			    End If
			  'Repetir Numeração de Peg em outra Filial
			  ElseIf qSamParametrosProcContas.FieldByName("REPETICAONUMERACAOPEG").Value = 4 Then
				bsShowMessage("Já existe um PEG com o mesmo número de Carta Remessa desse mesmo Recebedor nesta Filial!", "E")
			    CanContinue = False
			    Exit Sub
			  'Nunca Repetir Numeração de Peg
			  Else
				bsShowMessage("Já existe um PEG com o mesmo número de Carta Remessa desse mesmo Recebedor!", "E")
                CanContinue = False
                Exit Sub
			  End If
			End If
			Set CARTA2 = Nothing
		  End If
		End If
	  End If
  End If
  FinalizaSamParametrosProcContas

  InicializaSamParametrosProcContas("ACEITARECEBIMENTORETROATIVO")
  If qSamParametrosProcContas.FieldByName("ACEITARECEBIMENTORETROATIVO").AsString = "N" Then
	If CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime < Date Then
	  bsShowMessage("Data de Recebimento do PEG não pode ser 'MENOR' que a data corrente !", "E")
      CanContinue = False
      Exit Sub
	End If
  End If
  FinalizaSamParametrosProcContas

  InicializaSamParametrosProcContas("QTDANEXOSOBRIGATORIO")
  If qSamParametrosProcContas.FieldByName("QTDANEXOSOBRIGATORIO").AsString = "S" Then
	If CurrentQuery.FieldByName("QTDANEXOS").IsNull And CurrentQuery.FieldByName("TABREGIMEPGTO").Value = 2 Then
	  bsShowMessage("Campo Quantidade de Anexos é obrigatório para Reembolso.", "E")
      CanContinue = False
      Exit Sub
	End If
  End If
  FinalizaSamParametrosProcContas

  ' Verifica se a data de recebimento do PEG está válida
  If CurrentQuery.State <> 1 Then
	If Not VerificaDataRecebimento Then
	  CanContinue = False
	  bsShowMessage("Data Recebimento não pode ser menor que a data atendimento de alguma guia do PEG", "E")
	  Exit Sub
	End If
  End If

  Dim Dt As Date

  Dt = CurrentQuery.FieldByName("DATAEMISSAORECIBO").AsDateTime

  If Year(Dt) = 1899 Then
	If Month(Dt) = 12 Then
	  If Day(Dt) < 30 Then
		bsShowMessage("Data Inválida", "E")
		CanContinue = False
		DATAEMISSAORECIBO.SetFocus
		Exit Sub
	  End If
	Else
	  bsShowMessage("Data Inválida", "E")
	  CanContinue = False
	  DATAEMISSAORECIBO.SetFocus
	  Exit Sub
	End If
  ElseIf Year(Dt) < 1899 Then
	bsShowMessage("Data Inválida", "E")
	CanContinue = False
	DATAEMISSAORECIBO.SetFocus
	Exit Sub
  ElseIf Year(Dt) = 2099 Then
	If Month(Dt) = 12 Then
	  If Day(Dt) > 30 Then
		bsShowMessage("Data Inválida", "E")
		CanContinue = False
		DATAEMISSAORECIBO.SetFocus
		Exit Sub
	  End If
	End If
  End If

  If CurrentQuery.FieldByName("DATAEMISSAORECIBO").AsDateTime > ServerDate Then
	bsShowMessage("Entre com Data Emissão de Recibo menor ou igual a de hoje", "E")
	DATAEMISSAORECIBO.SetFocus
	CanContinue = False
	Exit Sub
  End If

  If CurrentQuery.State <> 1 Then
	If CurrentQuery.FieldByName("FILIALPROCESSAMENTO").IsNull Then
	  bsShowMessage("Filial processamento não pode estar nula.", "E")
	  CanContinue = False
	  Exit Sub
	End If
  End If

  If(CurrentQuery.State =2)Or(CurrentQuery.State =3)Then
	If (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime) Or _
	   (CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime < ServerDate) Then
	  bsShowMessage("A data de pagamento deve ser maior que a de recebimento e maior ou igual a de hoje", "E")
	  CanContinue = False
	  Exit Sub
	End If
  End If

  If CurrentQuery.State = 3 Then 'Somente ao inserir
	CurrentQuery.FieldByName("DATAPAGAMENTO").Value = Int(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)
  End If

  'Valida CPF/CNPJ quando for informada a conta de terceiros.
  If CurrentQuery.FieldByName("TABREGIMEPGTO").AsString = "2" Then
	If Len(CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString) <> 0 Then
	  If Not CheckCPFCNPJ(CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString, 0, True, Msg) Then
		bsShowMessage("O campo CPF/CNPJ da conta de terceiros é inválido", "E")
		CanContinue = False
		Exit Sub
	  End If
	End If
  End If

  If (EhPrestadorFixo(CurrentQuery.FieldByName("RECEBEDOR").AsInteger) And (CurrentQuery.FieldByName("VALORPAGAMENTORATEIO").IsNull)) Then
    VALORPAGAMENTORATEIO.SetFocus
    MsgBox("Recebedor informado é do tipo recebedor fixo. Informe o valor para rateio")
    CanContinue = False
    Exit Sub
  End If


  '********************************************************************************************************************************
  '********************************************************************************************************************************
  'Todas as verificações para não permitir salvar o PEG devem ser realizadas antes deste ponto
  '********************************************************************************************************************************
  '********************************************************************************************************************************


  'Quando alterar os dados da conta corrente no peg de reembolso, deverão ser automaticamente alterados os dados da conta
  'corrente das guias pertencentes ao mesmo.

  Dim qContaCorrente As Object
  Dim vDadosAtuais As String
  Dim qBuscaContaCorrente As Object

  'Verifica o valor dos respectivos campos na exibicao da interface. É feito a comparacao no momento de salvar o
  'PEG. Se ao salvar, algum desses campos estiverem com valor diferente de quando foi exibida a interface, o usuario
  'será questionado se deseja que os campos do conta corrente sejam repassados para a guia.

  Set qContaCorrente = NewQuery

  If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2) And _
	 (CurrentQuery.State = 2) Then

    If CurrentQuery.FieldByName("CREDITOCONTATERCEIROS").AsString = "N" Then
      CurrentQuery.FieldByName("BANCO").AsString = ""
      CurrentQuery.FieldByName("AGENCIA").AsString = ""
      CurrentQuery.FieldByName("CONTACORRENTENUMERO").AsString = ""		' apaga os campos referentes a credito de conta de terceiros,
      CurrentQuery.FieldByName("CONTACORRENTEDV").AsString = ""			' caso o flag seja desmarcado
      CurrentQuery.FieldByName("CONTACORRENTENOME").AsString = ""
      CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString = ""
    End If

	Set qBuscaContaCorrente = NewQuery

	qBuscaContaCorrente.Clear
	qBuscaContaCorrente.Add("SELECT BANCO,              ")
	qBuscaContaCorrente.Add("       AGENCIA,            ")
	qBuscaContaCorrente.Add("       CONTACORRENTENUMERO,")
	qBuscaContaCorrente.Add("       CONTACORRENTEDV,    ")
	qBuscaContaCorrente.Add("       CONTACORRENTENOME,  ")
	qBuscaContaCorrente.Add("       CONTACORRENTECPFCNPJ")
	qBuscaContaCorrente.Add("  FROM SAM_PEG             ")
	qBuscaContaCorrente.Add(" WHERE HANDLE =:HANDLEPEG  ")
	qBuscaContaCorrente.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qBuscaContaCorrente.Active = True

	'Verifica se houve alteracoes dos dados salvos anteriormente em relacao a digitacao atual.
	If ((qBuscaContaCorrente.FieldByName("BANCO").AsString) <> (CurrentQuery.FieldByName("BANCO").AsString)) Or _
	   ((qBuscaContaCorrente.FieldByName("AGENCIA").AsString) <> (CurrentQuery.FieldByName("AGENCIA").AsString)) Or _
	   ((qBuscaContaCorrente.FieldByName("CONTACORRENTENUMERO").AsString) <> (CurrentQuery.FieldByName("CONTACORRENTENUMERO").AsString)) Or _
	   ((qBuscaContaCorrente.FieldByName("CONTACORRENTEDV").AsString) <> (CurrentQuery.FieldByName("CONTACORRENTEDV").AsString)) Or _
	   ((qBuscaContaCorrente.FieldByName("CONTACORRENTENOME").AsString) <> (CurrentQuery.FieldByName("CONTACORRENTENOME").AsString)) Or _
	   ((qBuscaContaCorrente.FieldByName("CONTACORRENTECPFCNPJ").AsString) <> (CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString)) Then
	  Set qBuscaContaCorrente = NewQuery

	  qContaCorrente.Clear
	  qContaCorrente.Add("UPDATE SAM_GUIA                                    ")
	  qContaCorrente.Add("   SET BANCO                = :BANCO,              ")
	  qContaCorrente.Add("       AGENCIA              = :AGENCIA,            ")
	  qContaCorrente.Add("       CONTACORRENTENUMERO  = :CONTACORRENTENUMERO,")
	  qContaCorrente.Add("       CONTACORRENTEDV      = :CONTACORRENTEDV,    ")
	  qContaCorrente.Add("       CONTACORRENTENOME    = :CONTACORRENTENOME,  ")
	  qContaCorrente.Add("       CONTACORRENTECPFCNPJ = :CONTACORRENTECPFCNPJ")
	  qContaCorrente.Add(" WHERE PEG                  = :HANDLEPEG           ")

      'Se o banco for nulo nao faz a exportacao do campo do peg para a guia
	  If CurrentQuery.FieldByName("BANCO").IsNull Then
		qContaCorrente.ParamByName("BANCO").DataType = ftInteger
		qContaCorrente.ParamByName("BANCO").Clear
	  Else
		qContaCorrente.ParamByName("BANCO").Value = CurrentQuery.FieldByName("BANCO").Value
	  End If

	  'Se a agencia for nula nao faz a exportacao do campo do peg para a guia
	  If CurrentQuery.FieldByName("AGENCIA").IsNull Then
		qContaCorrente.ParamByName("AGENCIA").DataType = ftInteger
		qContaCorrente.ParamByName("AGENCIA").Clear
	  Else
		qContaCorrente.ParamByName("AGENCIA").Value = CurrentQuery.FieldByName("AGENCIA").Value
	  End If

	  'Se a conta corrente for nula nao faz a exportacao do campo do peg para a guia
	  If CurrentQuery.FieldByName("CONTACORRENTENUMERO").IsNull Then
		qContaCorrente.ParamByName("CONTACORRENTENUMERO").DataType = ftInteger
		qContaCorrente.ParamByName("CONTACORRENTENUMERO").Clear
	  Else
		qContaCorrente.ParamByName("CONTACORRENTENUMERO").Value = CurrentQuery.FieldByName("CONTACORRENTENUMERO").Value
	  End If

	  'Se o digito da conta corrente for nulo nao faz a exportacao do campo do peg para a guia
	  If CurrentQuery.FieldByName("CONTACORRENTEDV").IsNull Then
		qContaCorrente.ParamByName("CONTACORRENTEDV").DataType = ftInteger
		qContaCorrente.ParamByName("CONTACORRENTEDV").Clear
	  Else
		qContaCorrente.ParamByName("CONTACORRENTEDV").Value = CurrentQuery.FieldByName("CONTACORRENTEDV").Value
	  End If

	  'Se o nome do titular da conta corrente for nulo nao faz a exportacao do campo do peg para a guia
	  If CurrentQuery.FieldByName("CONTACORRENTENOME").IsNull Then
		qContaCorrente.ParamByName("CONTACORRENTENOME").DataType = ftInteger
		qContaCorrente.ParamByName("CONTACORRENTENOME").Clear
	  Else
		qContaCorrente.ParamByName("CONTACORRENTENOME").Value = CurrentQuery.FieldByName("CONTACORRENTENOME").Value
	  End If

	  'Se o cpf/cnpj do titular da conta corrente for nulo nao faz a exportacao do campo do peg para a guia
	  If CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").IsNull Then
		qContaCorrente.ParamByName("CONTACORRENTECPFCNPJ").DataType = ftInteger
		qContaCorrente.ParamByName("CONTACORRENTECPFCNPJ").Clear
	  Else
		qContaCorrente.ParamByName("CONTACORRENTECPFCNPJ").Value = CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").Value
	  End If

	  qContaCorrente.ParamByName("HANDLEPEG").Value = CurrentQuery.FieldByName("HANDLE").Value
	  qContaCorrente.ExecSQL

	  Set qContaCorrente = Nothing
	End If
  End If

  Set qBuscaContaCorrente = Nothing

  If viHTipoPegAnterior <> CurrentQuery.FieldByName("TIPOPEG").AsInteger Then
	If PegReembolsoMatMed Then
	  bsShowMessage("Não é possível alterar o tipo de PEG envolvendo Reembolso de MatMed com Guias já cadastradas!", "E")
	  CanContinue = False
	  Exit Sub
	End If
	CurrentQuery.FieldByName("TIPOPEGANTERIOR").Clear
  End If

  If (CurrentQuery.State = 3) Then
	CurrentQuery.FieldByName("DIGITADOR").AsInteger = CurrentUser
	CurrentQuery.FieldByName("DATADIGITACAO").AsDateTime = ServerNow
  End If

  If CurrentQuery.State = 3 Then
    Dim vsRegraCalcBaseIss As String
    vsRegraCalcBaseIss = IIf(Len(CurrentQuery.FieldByName("REGRACALCBASEISS").AsString) > 0, CurrentQuery.FieldByName("REGRACALCBASEISS").AsString, "N")

    InicializaSamParametrosProcContas("CALCISSVALORBRUTOLIQUIDOPEG")
    If qSamParametrosProcContas.FieldByName("CALCISSVALORBRUTOLIQUIDOPEG").Value = "S" And Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
      Set qRegraBaseISS = NewQuery

      qRegraBaseISS.Clear
      qRegraBaseISS.Add("SELECT REGRACALCBASEISS ")
      qRegraBaseISS.Add("  FROM SAM_PRESTADOR    ")
      qRegraBaseISS.Add(" WHERE HANDLE = :HANDLE ")
      qRegraBaseISS.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
      qRegraBaseISS.Active = True

      If Not qRegraBaseISS.FieldByName("REGRACALCBASEISS").IsNull Then
        CurrentQuery.FieldByName("REGRACALCBASEISS").Value = qRegraBaseISS.FieldByName("REGRACALCBASEISS").Value
      Else
        CurrentQuery.FieldByName("REGRACALCBASEISS").Value = "N"
      End If
    Else
      CurrentQuery.FieldByName("REGRACALCBASEISS").Value = "N"
    End If
    FinalizaSamParametrosProcContas

	If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) Then
	  If vsRegraCalcBaseIss <> CurrentQuery.FieldByName("REGRACALCBASEISS").AsString Then
	    If CurrentQuery.FieldByName("REGRACALCBASEISS").Value = "N" Then
	      bsShowMessage("Regra p/ cálculo de base ISS alterado para 'Normal' de acordo com o parâmetro geral e cadastro do recebedor.", "I")
	    ElseIf CurrentQuery.FieldByName("REGRACALCBASEISS").Value = "B" Then
	      bsShowMessage("Regra p/ cálculo de base ISS alterado para 'Valor Bruto' de acordo com o parâmetro geral e cadastro do recebedor.", "I")
	    ElseIf CurrentQuery.FieldByName("REGRACALCBASEISS").Value = "L" Then
	      bsShowMessage("Regra p/ cálculo de base ISS alterado para 'Valor Líquido' de acordo com o parâmetro geral e cadastro do recebedor.", "I")
	    End If
	  End If
	End If

	If CurrentQuery.FieldByName("REGRACALCBASECONTFEDERAIS").IsNull Then

      InicializaSamParametrosProcContas("CALCCONTFEDERAISBRUTOLIQPEG")
      If qSamParametrosProcContas.FieldByName("CALCCONTFEDERAISBRUTOLIQPEG").Value = "S" And Not CurrentQuery.FieldByName("RECEBEDOR").IsNull Then

        Set qRegraBaseContribFederais = NewQuery
        qRegraBaseContribFederais.Clear
        qRegraBaseContribFederais.Add("SELECT REGRACALCBASECONTFEDERAIS ")
        qRegraBaseContribFederais.Add("  FROM SAM_PRESTADOR    ")
        qRegraBaseContribFederais.Add(" WHERE HANDLE = :HANDLE ")
        qRegraBaseContribFederais.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
        qRegraBaseContribFederais.Active = True

        If Not qRegraBaseContribFederais.FieldByName("REGRACALCBASECONTFEDERAIS").IsNull Then
          CurrentQuery.FieldByName("REGRACALCBASECONTFEDERAIS").Value = qRegraBaseContribFederais.FieldByName("REGRACALCBASECONTFEDERAIS").Value
        Else
          CurrentQuery.FieldByName("REGRACALCBASECONTFEDERAIS").Value = "N"
        End If
      Else
        CurrentQuery.FieldByName("REGRACALCBASECONTFEDERAIS").Value = "N"
      End If
      FinalizaSamParametrosProcContas

	End If
  End If

  Set qVerificaConsiderarSp = NewQuery
  Set qServicosPrestador = NewQuery

  qVerificaConsiderarSp.Active = False
  qVerificaConsiderarSp.Clear
  qVerificaConsiderarSp.Add("SELECT CONSIDERARCODSERVICO FROM SFN_PARAMETROSFIN")
  qVerificaConsiderarSp.Active = True

  qServicosPrestador.Active = False
  qServicosPrestador.Clear
  qServicosPrestador.Add("SELECT * FROM SAM_PRESTADOR WHERE HANDLE = :PPRESTHANDLE")
  qServicosPrestador.ParamByName("PPRESTHANDLE").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  qServicosPrestador.Active = True

  If (qVerificaConsiderarSp.FieldByName("CONSIDERARCODSERVICO").AsString = "S") Then
    If (CurrentQuery.FieldByName("TABREGIMEPGTO").Value = "1") Then
      If ((CurrentQuery.FieldByName("LISTASERVICO").IsNull) And (CurrentQuery.FieldByName("CODIGOSERVICO").IsNull)) Then
        If ((Not qServicosPrestador.FieldByName("CODIGOSERVICO").IsNull) And (Not qServicosPrestador.FieldByName("CODIGOSERVICOPREFSP").IsNull)) Then
          CurrentQuery.FieldByName("LISTASERVICO").AsInteger = qServicosPrestador.FieldByName("CODIGOSERVICO").AsInteger
          CurrentQuery.FieldByName("CODIGOSERVICO").AsInteger = qServicosPrestador.FieldByName("CODIGOSERVICOPREFSP").AsInteger
        Else
          bsshowmessage("Necessário informar um serviço no Prestador Recebedor ou no PEG","E")
          CanContinue = False
          Exit Sub
        End If
      End If
    End If
  End If

  If ((CurrentQuery.FieldByName("LISTASERVICO").IsNull) And (Not CurrentQuery.FieldByName("CODIGOSERVICO").IsNull)) Then
    bsshowmessage("Necessário informar um Código de Serviço referente ao Municípo de São Paulo","E")
    CanContinue = False
    Exit Sub
  End If

  If(Not (CurrentQuery.FieldByName("CODBARRASBOLETO").IsNull) And (CurrentQuery.FieldByName("TOTALPAGARINFORMADO").IsNull) And CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "S")Then
     bsshowmessage("Necessário informar o Total Previsto a Pagar quando o PEG possuir informação de Código de Barras","E")
     CanContinue = False
     Exit Sub
  End If


  Set qServicosPrestador = Nothing
  Set qVerificaConsiderarSp = Nothing
  Set qRegraBaseISS = Nothing
  Set qRegraBaseContribFederais = Nothing


  If WebMode And (OLDRECEBEDOR <> CurrentQuery.FieldByName("RECEBEDOR").AsInteger Or CurrentQuery.FieldByName("PERCENTUALREDUCAOINSS").IsNull) Then
	AtualizarDescontoINSS
  End If

End Sub


Public Sub TABLE_NewRecord()
  Dim vRegime As Integer
  CurrentQuery.FieldByName("DATA").Value = ServerDate
  ' Campos variaveis carregados no before insert
  If vgFilialProcessamento > 0 Then
	CurrentQuery.FieldByName("FILIALPROCESSAMENTO").AsInteger =vgFilialProcessamento
  End If

  If vgFilial > 0 Then
	CurrentQuery.FieldByName("FILIAL").AsInteger = vgFilial
  End If

  InicializaSamParametrosProcContas("SUGEREPEGCARTAREMESSA, SEMPREHERDARPEGDACARTAREMESSA")
  If (qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean) Or (qSamParametrosProcContas.FieldByName("SEMPREHERDARPEGDACARTAREMESSA").AsBoolean) Then
	CurrentQuery.FieldByName("CARTAREMESSA").Value = CurrentQuery.FieldByName("PEG").Value
  End If
  FinalizaSamParametrosProcContas

  RECEBEDOR.ReadOnly = AtribuirReadOnly(False)
  LOCALEXECUCAO.ReadOnly = AtribuirReadOnly(False)
  BENEFICIARIO.ReadOnly = AtribuirReadOnly(False)
  vSituacaoAnteriorPeg = CurrentQuery.FieldByName("SITUACAO").AsString
End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "CONFERIDO"
			CONFERIDO_OnClick
		Case "DESDOBRAR"
			DESDOBRAR_OnClick
		Case "DEVOLVERPEG"
			DEVOLVERPEG_OnClick
		Case "REPROCESSARPEG"
			REPROCESSARPEG_OnClick
		Case "FASEPEG"
			FASEPEG_OnClick
		Case "CRITICARDIGITACAO"
		    CRITICARDIGITACAO_OnClick
		Case "BOTAORECLASSIFICAR"
		    BOTAORECLASSIFICAR_OnClick
		Case "EXCLUIRPEG"
		    EXCLUIRPEG_OnClick
		Case "BOTAOCANCELARPROVISAO"
			BOTAOCANCELARPROVISAO_OnClick
		Case "BOTAOVERIFICAMONITORAMENTO"
		  IncluiSessionVarMonitoramento
		Case "BOTAOLIBERARVERIFICACAO"
		  BOTAOLIBERARVERIFICACAO_OnClick
		Case "BOTAOASSUMIRANALISE"
		  BOTAOASSUMIRANALISE_OnClick
		Case "BOTAOTRIAGEM"
		  BOTAOENCAMINHAR_OnClick
		Case "BOTAOPROVISIONARPEG"
		  BOTAOPROVISIONARPEG_OnClick
		Case "BOTAOGLOSATOTAL"
		  BOTAOGLOSATOTAL_OnClick
		Case "BOTAOINCLUIRPRESTADOR"
		  BOTAOINCLUIRPRESTADOR_OnClick
		Case "BOTAOVERIFICAMONITORAMENTO"
		  BOTAOVERIFICAMONITORAMENTO_OnClick
		Case "BOTAOALTERARDADOSNF"
		  BOTAOALTERARDADOSNF_OnClick
		Case "BOTAOALTERARDATACONTABIL"
		  BOTAOALTERARDATACONTABIL_OnClick
		Case "BOTAOALTERARDATAPAGAMENTO"
		  BOTAOALTERARDATAPAGAMENTO_OnClick
		Case "BOTAOALTERARVALORAPRESENTADO"
		  BOTAOALTERARVALORAPRESENTADO_OnClick
		Case "BOTAOALTERARGUIASAPRESENTADAS"
		  BOTAOALTERARGUIASAPRESENTADAS_OnClick
	End Select
End Sub


Public Sub TABLE_UpdateRequired()
  If WebMode Then
    If Not CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
      DATAPAGAMENTO_OnExit
    End If
    If CurrentQuery.FieldByName("NUMEROPAGAMENTO").IsNull Or CurrentQuery.FieldByName("DATAPAGAMENTO").IsNull Then
      VERIFICAPAG("")
    End If

    'simular na web o sugerir numero do peg e carta remessa
    If Not CurrentQuery.FieldByName("CARTAREMESSA").IsNull Then
      CARTAREMESSA_OnExit
    Else
      If Not CurrentQuery.FieldByName("PEG").IsNull Then
        PEG_OnExit
      End If
    End If
  End If
End Sub



Public Sub TABREGIMEPGTO_OnChange()
  vMudouTabRegimePagto = True

  If TABREGIMEPGTO.PageIndex = 0 Then
    VTabRegimePgtoAnterior = 1
  Else
    VTabRegimePgtoAnterior = 0
  End If

  If TABREGIMEPGTO.PageIndex = 1 Then 'reembolso
	CurrentQuery.FieldByName("CREDITOCONTATERCEIROS").AsString = "N"  'Vieira - SMS 68005 - 07/09/2006

	GRUPOADIANTAMENTO.Visible = False

	If VerificaRegimePgto = 1 Then
	  bsShowMessage("Regime de pagamento não compatível com tipo de peg", "I")
	  CurrentQuery.FieldByName("TABREGIMEPGTO").Value = VerificaRegimePgto
	End If

    InicializaSamParametrosProcContas("ORIGEMREEMBOLSO")
	If qSamParametrosProcContas.FieldByName("ORIGEMREEMBOLSO").AsString = "2" Then 'Executor
	  ESTADO.ReadOnly = AtribuirReadOnly(True) 'sms 63627 - Edilson.Castro - 14/06/2006
	  MUNICIPIO.ReadOnly = AtribuirReadOnly(True) 'sms 63627 - Edilson.Castro - 14/06/2006
	  CurrentQuery.FieldByName("ESTADO").Clear
	  CurrentQuery.FieldByName("MUNICIPIO").Clear
	Else
	  ESTADO.ReadOnly = AtribuirReadOnly(False) 'sms 63627 - Edilson.Castro - 14/06/2006
	  MUNICIPIO.ReadOnly = AtribuirReadOnly(False) 'sms 63627 - Edilson.Castro - 14/06/2006
	End If
	FinalizaSamParametrosProcContas

  Else
	GRUPOADIANTAMENTO.Visible = True

	If VerificaRegimePgto = 2 Then
	  bsShowMessage("Regime de pagamento não compatível com tipo de peg", "E")
	  CurrentQuery.FieldByName("TABREGIMEPGTO").Value = VerificaRegimePgto
	End If
  End If
End Sub


Public Sub TABREGIMEPGTO_OnChanging(AllowChange As Boolean)
  If CurrentQuery.State <> 3 Then 'diferente de Inclusão
	bsShowMessage("Alteração não permitida", "I")
	AllowChange = False
  End If
End Sub

Public Sub TOTALPAGARINFORMADO_OnExit()
  If CurrentQuery.State = 1 Then
	Exit Sub
  End If

  processachange = False

  Dim A As Double
  Dim b As Double
  Dim c As Long
  Dim d As String
  Dim e As String
  Dim f As String

  If CurrentQuery.State = 2 Or CurrentQuery.State = 3 Then
	If TemRegra(A, b, c, d, e, f) And CurrentQuery.FieldByName("TOTALPAGARINFORMADO").AsFloat > 0 Then
	  If CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "N" Then
		bsShowMessage("Existe regra de adiantamento para este Recebedor", "I")
	  End If

	  CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "S"
	Else
	  CurrentQuery.FieldByName("ADIANTAMENTO").AsString = "N"
	End If
  End If

  CalculaAdiantamento
End Sub


Public Sub VALORACRESCIMOISSINFORMADO_OnExit()
  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
	CurrentQuery.FieldByName("VALORACRESCIMOISS").AsFloat = CurrentQuery.FieldByName("VALORACRESCIMOISSINFORMADO").AsFloat
  End If
End Sub


Public Sub VALORADIANTAMENTO_OnExit()
  If VALORADIANTAMENTO.ReadOnly = False Then
    gchangeValorAdi = False
    VerificaAdiantamento False, True
  End If
End Sub


Public Sub VALORDESCONTO_OnExit()
  If VALORDESCONTO.ReadOnly = False Then
    gchangeValorAdi =False
	VerificaAdiantamento True, False
  End If
End Sub


Public Sub QuatidadeGuia
  Dim Qguia As Object
  Set Qguia = NewQuery

  Qguia.Add("SELECT COUNT(1) QTD FROM SAM_GUIA WHERE PEG = :peg")

  Qguia.ParamByName("PEG").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Qguia.Active = True

  QTDGUIADIGITADA.Text = "Digitada: " + Qguia.FieldByName("QTD").AsString
  DIFERENCADEGUIAS.Text = "Diferença: " + Str(CurrentQuery.FieldByName("QTDGUIA").AsInteger - Qguia.FieldByName("QTD").AsString - CurrentQuery.FieldByName("QTDGUIADEVOLVIDA").AsInteger)

  Set Qguia = Nothing
End Sub


Public Function VerificaDataRecebimento As Boolean
  'Verifica se a data de recebimento do PEG é maior ou igual que a maior data de atendimento das guias do PEG
  Dim q As Object
  Set q =NewQuery
  q.Clear
  q.Add("SELECT MAX(DATAATENDIMENTO) DATAATENDIMENTO")
  q.Add("  FROM SAM_GUIA                            ")
  q.Add(" WHERE PEG = :PEG                          ")
  q.Add("   AND DATAATENDIMENTO IS NOT NULL         ")
  q.ParamByName("PEG").AsInteger =CurrentQuery.FieldByName("HANDLE").AsInteger
  q.Active =True
  If Not q.EOF Then
    If Int(CurrentQuery.FieldByName("DATARECEBIMENTO").AsDateTime) >= Int(q.FieldByName("DATAATENDIMENTO").AsDateTime) Then
      VerificaDataRecebimento =True
    Else
      VerificaDataRecebimento =False
    End If
  Else
    VerificaDataRecebimento =True
  End If
  Set q =Nothing
End Function


Public Function VerificaRegimePgto As Integer
  Dim qTipoPeg As Object
  If RecordHandleOfTableInterfacePEG("SAM_TIPOPEG")>0 Then
    Set qTipoPeg =NewQuery
    qTipoPeg.Active =False
    qTipoPeg.Clear
    qTipoPeg.Add("SELECT TABREGIMEPGTO")
    qTipoPeg.Add("  FROM SAM_TIPOPEG")
    qTipoPeg.Add(" WHERE HANDLE=:TIPOPEG")
    qTipoPeg.ParamByName("TIPOPEG").Value =RecordHandleOfTableInterfacePEG("SAM_TIPOPEG")
    qTipoPeg.Active =True
    VerificaRegimePgto =qTipoPeg.FieldByName("TABREGIMEPGTO").Value
  End If
End Function

Function PermitePeloGrupoSeguranca(pTabela, pCampo As String) As Boolean
  Dim viHandleGrupoTabelaCampo As Long
  Dim qVerificaPermissaoGrupoSeguranca As Object
  Set qVerificaPermissaoGrupoSeguranca = NewQuery

  qVerificaPermissaoGrupoSeguranca.Clear
  qVerificaPermissaoGrupoSeguranca.Add("SELECT COUNT(1) QTDE")
  qVerificaPermissaoGrupoSeguranca.Add("  FROM Z_GRUPOUSUARIOS      GU,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELAS       GT, ")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_TABELAS             T,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELACAMPOS GTC,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_CAMPOS              C")
  qVerificaPermissaoGrupoSeguranca.Add(" WHERE GU.GRUPO = GT.GRUPO")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GT.TABELA = T.HANDLE")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.GRUPOTABELA = GT.HANDLE")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.CAMPO = C.HANDLE")
  qVerificaPermissaoGrupoSeguranca.Add("   AND T.NOME = :TABELA")
  qVerificaPermissaoGrupoSeguranca.Add("   AND C.NOME = :CAMPO")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GU.HANDLE = :HUSUARIO")
  qVerificaPermissaoGrupoSeguranca.ParamByName("TABELA").AsString = pTabela
  qVerificaPermissaoGrupoSeguranca.ParamByName("CAMPO").AsString = pCampo
  qVerificaPermissaoGrupoSeguranca.ParamByName("HUSUARIO").AsInteger = CurrentUser
  qVerificaPermissaoGrupoSeguranca.Active = True

  If qVerificaPermissaoGrupoSeguranca.FieldByName("QTDE").AsInteger = 0 Then
    PermitePeloGrupoSeguranca = True
    Set qVerificaPermissaoGrupoSeguranca = Nothing
    Exit Function
  Else
    qVerificaPermissaoGrupoSeguranca.Active = False
    qVerificaPermissaoGrupoSeguranca.Clear
    qVerificaPermissaoGrupoSeguranca.Add("SELECT COUNT(1) QTDE")
    qVerificaPermissaoGrupoSeguranca.Add("  FROM Z_GRUPOUSUARIOGRUPOS GU,")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELAS       GT, ")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_TABELAS             T,")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELACAMPOS GTC,")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_CAMPOS              C")
    qVerificaPermissaoGrupoSeguranca.Add(" WHERE GU.GRUPOADICIONADO = GT.GRUPO")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GT.TABELA = T.HANDLE")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.GRUPOTABELA = GT.HANDLE")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.CAMPO = C.HANDLE")
    qVerificaPermissaoGrupoSeguranca.Add("   AND T.NOME = :TABELA")
    qVerificaPermissaoGrupoSeguranca.Add("   AND C.NOME = :CAMPO")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GU.USUARIO = :HUSUARIO")
    qVerificaPermissaoGrupoSeguranca.ParamByName("TABELA").AsString = pTabela
    qVerificaPermissaoGrupoSeguranca.ParamByName("CAMPO").AsString = pCampo
    qVerificaPermissaoGrupoSeguranca.ParamByName("HUSUARIO").AsInteger = CurrentUser
    qVerificaPermissaoGrupoSeguranca.Active = True
    If qVerificaPermissaoGrupoSeguranca.FieldByName("QTDE").AsInteger = 0 Then
      PermitePeloGrupoSeguranca = True
      Set qVerificaPermissaoGrupoSeguranca = Nothing
      Exit Function
    End If
  End If

  qVerificaPermissaoGrupoSeguranca.Clear
  qVerificaPermissaoGrupoSeguranca.Add("SELECT COUNT(1) QTDE")
  qVerificaPermissaoGrupoSeguranca.Add("  FROM Z_GRUPOUSUARIOS      GU,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELAS       GT,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_TABELAS             T,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELACAMPOS GTC,")
  qVerificaPermissaoGrupoSeguranca.Add("       Z_CAMPOS              C")
  qVerificaPermissaoGrupoSeguranca.Add(" WHERE GU.GRUPO = GT.GRUPO")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GT.TABELA = T.HANDLE")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.GRUPOTABELA = GT.HANDLE")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.CAMPO = C.HANDLE")
  qVerificaPermissaoGrupoSeguranca.Add("   AND T.NOME = :TABELA")
  qVerificaPermissaoGrupoSeguranca.Add("   AND C.NOME = :CAMPO")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GU.HANDLE = :HUSUARIO")
  qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.ALTERAR = 'S'")
  qVerificaPermissaoGrupoSeguranca.ParamByName("TABELA").AsString = pTabela
  qVerificaPermissaoGrupoSeguranca.ParamByName("CAMPO").AsString = pCampo
  qVerificaPermissaoGrupoSeguranca.ParamByName("HUSUARIO").AsInteger = CurrentUser
  qVerificaPermissaoGrupoSeguranca.Active = True

  If qVerificaPermissaoGrupoSeguranca.FieldByName("QTDE").AsInteger > 0 Then 'Permite executar
    PermitePeloGrupoSeguranca = True
  Else 'Vai verificar se pode executar pelos grupos adicionais
    qVerificaPermissaoGrupoSeguranca.Active = False
    qVerificaPermissaoGrupoSeguranca.Clear
    qVerificaPermissaoGrupoSeguranca.Add("SELECT COUNT(1) QTDE")
    qVerificaPermissaoGrupoSeguranca.Add("  FROM Z_GRUPOUSUARIOGRUPOS GUG,")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELAS        GT,")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_GRUPOTABELACAMPOS  GTC,")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_TABELAS              T,       ")
    qVerificaPermissaoGrupoSeguranca.Add("       Z_CAMPOS               C")
    qVerificaPermissaoGrupoSeguranca.Add(" WHERE GUG.GRUPOADICIONADO = GT.GRUPO")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.GRUPOTABELA = GT.HANDLE")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GT.TABELA = T.HANDLE")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.CAMPO = C.HANDLE")
    qVerificaPermissaoGrupoSeguranca.Add("   AND T.NOME = :TABELA")
    qVerificaPermissaoGrupoSeguranca.Add("   AND C.NOME = :CAMPO")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GUG.USUARIO = :HUSUARIO")
    qVerificaPermissaoGrupoSeguranca.Add("   AND GTC.ALTERAR = 'S'")
    qVerificaPermissaoGrupoSeguranca.ParamByName("TABELA").AsString = pTabela
    qVerificaPermissaoGrupoSeguranca.ParamByName("CAMPO").AsString = pCampo
    qVerificaPermissaoGrupoSeguranca.ParamByName("HUSUARIO").AsInteger = CurrentUser
    qVerificaPermissaoGrupoSeguranca.Active = True

    If qVerificaPermissaoGrupoSeguranca.FieldByName("QTDE").AsInteger = 0 Then 'Não pode executar
      PermitePeloGrupoSeguranca = False
    Else 'Pode executar
      PermitePeloGrupoSeguranca = True
    End If
  End If

  Set qVerificaPermissaoGrupoSeguranca = Nothing
End Function


Public Function Arredonda(Valor As Double) As Double
  Arredonda=CCur(Format(Valor,"###,###,##0.00"))
End Function


Public Function EhPrestadorFixo(pHandlePrestador As Long) As Boolean
  Dim qTIPOPRESTADOR As Object
  Set qTIPOPRESTADOR = NewQuery

  qTIPOPRESTADOR.Clear
  qTIPOPRESTADOR.Add("SELECT TABTIPO TIPOPRESTADOR  ")
  qTIPOPRESTADOR.Add("  FROM SAM_PRESTADOR_TIPOPAGTO")
  qTIPOPRESTADOR.Add(" WHERE PRESTADOR = :RECEBEDOR ")
  qTIPOPRESTADOR.ParamByName("RECEBEDOR").AsInteger = pHandlePrestador
  qTIPOPRESTADOR.Active = True

  EhPrestadorFixo = (qTIPOPRESTADOR.FieldByName("TIPOPRESTADOR").AsInteger = 2)
End Function


Public Sub VALORPAGAMENTORATEIO_OnExit()
  If ((VALORPAGAMENTORATEIO.ReadOnly) Or (CurrentQuery.State = 1)) Then
    Exit Sub
  End If

  Dim qTIPOPRESTADOR   As Object
  Dim qSALDOPRESTADOR  As Object
  Dim vrSaldoPrestador As Long
  Set qTIPOPRESTADOR  = NewQuery
  Set qSALDOPRESTADOR = NewQuery

  qTIPOPRESTADOR.Clear
  qTIPOPRESTADOR.Add("SELECT VALOR   VALORMAXIMO    ")
  qTIPOPRESTADOR.Add("  FROM SAM_PRESTADOR_TIPOPAGTO")
  qTIPOPRESTADOR.Add(" WHERE PRESTADOR = :RECEBEDOR ")
  qTIPOPRESTADOR.ParamByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  qTIPOPRESTADOR.Active = True

	qSALDOPRESTADOR.Clear
	qSALDOPRESTADOR.Add("SELECT SUM(VALOR) TOTALRECEBIDO          ")
	qSALDOPRESTADOR.Add("  FROM SAM_PRESTADOR_SALDOPAGAMENTO      ")
	qSALDOPRESTADOR.Add(" WHERE PRESTADOR = :RECEBEDOR            ")
	qSALDOPRESTADOR.Add("   AND DATAPAGAMENTO BETWEEN :DATAINICIAL")
	qSALDOPRESTADOR.Add("                         AND :DATAFINAL  ")
	qSALDOPRESTADOR.ParamByName("RECEBEDOR").AsInteger    = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
	qSALDOPRESTADOR.ParamByName("DATAINICIAL").AsDateTime = PrimeiroDiaCompetencia(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)
	qSALDOPRESTADOR.ParamByName("DATAFINAL").AsDateTime   = UltimoDiaCompetencia(CurrentQuery.FieldByName("DATAPAGAMENTO").AsDateTime)
	qSALDOPRESTADOR.Active = True

	vrSaldoPrestador = (qTIPOPRESTADOR.FieldByName("VALORMAXIMO").AsFloat - qSALDOPRESTADOR.FieldByName("TOTALRECEBIDO").AsFloat)

	If CurrentQuery.FieldByName("VALORPAGAMENTORATEIO").AsFloat > vrSaldoPrestador Then
	  MsgBox("Valor informado excede o valor fixado para o recebedor. Máximo disponível será utilizado")
      CurrentQuery.FieldByName("VALORPAGAMENTORATEIO").AsFloat = vrSaldoPrestador

	ElseIf CurrentQuery.FieldByName("VALORPAGAMENTORATEIO").IsNull Then
	  MsgBox("Recebedor informado é do tipo recebedor fixo. É necessário informar o valor para rateio")

	End If

  Set qTIPOPRESTADOR  = Nothing
  Set qSALDOPRESTADOR = Nothing
End Sub


Public Sub VOLTARSITUACAO_OnClick()
  Dim Interface As Object
  Dim Aux As Boolean
  Dim vFilial As Long
  If (VisibleMode Or WebMode) Then
    vFilial = RecordHandleOfTableInterfacePEG("FILIAIS")
  Else
    vFilial = CurrentQuery.FieldByName("FILIAL").AsInteger
  End If

  Aux = CheckFilialProcessamento(CurrentSystem, vFilial, "P")

  If Not Aux Then
    AtualizarCarga(True)
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SAMPEG.Processar")
  Interface.VoltarSituacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing

  AtualizarCarga(False)
End Sub


Public Function AtribuirReadOnly(Valor As Boolean) As Boolean
  'Esta função foi criada para evitar que campos fiquem habilitados em cargas read-only
  AtribuirReadOnly = IIf(CurrentQuery.RequestLive, Valor, True)
End Function


Public Function PegReembolsoMatMed As Boolean
  Dim qSAMTIPOPEG As Object
  Dim qSAMGUIA As Object
  Set qSAMTIPOPEG = NewQuery
  Set qSAMGUIA = NewQuery

  Dim vHANDLE As Long
  If (VisibleMode Or WebMode) Then
    vHANDLE = RecordHandleOfTableInterfacePEG("SAM_PEG")
  Else
    vHANDLE = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If

  qSAMTIPOPEG.Clear
  qSAMTIPOPEG.Add("SELECT T.REEMBOLSOMATMED ")
  qSAMTIPOPEG.Add("  FROM SAM_GUIA G        ")
  qSAMTIPOPEG.Add("  JOIN SAM_PEG  P ON (P.HANDLE = G.PEG)         ")
  qSAMTIPOPEG.Add("  JOIN SAM_TIPOPEG  T ON (T.HANDLE = P.TIPOPEG) ")
  qSAMTIPOPEG.Add(" WHERE P.HANDLE = :HANDLE")
  qSAMTIPOPEG.ParamByName("HANDLE").AsInteger = vHANDLE
  qSAMTIPOPEG.Active = True

  If Not qSAMTIPOPEG.EOF Then
    qSAMGUIA.Clear
    qSAMGUIA.Add("SELECT REEMBOLSOMATMED ")
    qSAMGUIA.Add("  FROM SAM_TIPOPEG     ")
    qSAMGUIA.Add(" WHERE HANDLE = :HANDLE")
    qSAMGUIA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOPEG").AsInteger
    qSAMGUIA.Active = True

    If ((qSAMGUIA.FieldByName("REEMBOLSOMATMED").AsString = "S") And (qSAMTIPOPEG.FieldByName("REEMBOLSOMATMED").AsString <> "S")) Or _
       ((qSAMGUIA.FieldByName("REEMBOLSOMATMED").AsString <> "S") And (qSAMTIPOPEG.FieldByName("REEMBOLSOMATMED").AsString = "S")) Then
      PegReembolsoMatMed = True
    Else
      PegReembolsoMatMed = False
    End If
  Else
    PegReembolsoMatMed = False
  End If

  Set qSAMTIPOPEG = Nothing
  Set qSAMGUIA = Nothing
End Function


Public Sub VerificaReciboNF()
  Dim qPrestador       As Object
  Dim vContadorRecibo  As Long

  If CurrentQuery.FieldByName("RECEBEDOR").IsNull Then
    RECIBO.ReadOnly = False
    NFNUMERO.ReadOnly = False
    DATAEMISSAONOTA.ReadOnly = False
  Else
    Set qPrestador       = NewQuery

    InicializaSamParametrosProcContas("MUDAFASEPEGCOMNFPEG")
    If qSamParametrosProcContas.FieldByName("MUDAFASEPEGCOMNFPEG").AsString = "S" Then
      With qPrestador
        .Active = False

        .Clear

        .Add("SELECT FISICAJURIDICA")
        .Add("  FROM SAM_PRESTADOR")
        .Add(" WHERE HANDLE = :RECEBEDOR")

        .ParamByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
        .Active = True
      End With

      If qPrestador.FieldByName("FISICAJURIDICA").AsInteger = 1 Then 'Física
        RECIBO.ReadOnly = False
        NFNUMERO.ReadOnly = True
        DATAEMISSAONOTA.ReadOnly = True

        If CurrentQuery.State <>1 Then
          If CurrentQuery.FieldByName("RECIBO").IsNull Then
            NewCounter("SAM_PEG_RECIBO", 0, 1, vContadorRecibo)

            If vContadorRecibo = 0 Then
              NewCounter("SAM_PEG_RECIBO", 0, 1, vContadorRecibo)
            End If

            CurrentQuery.FieldByName("RECIBO").AsInteger = vContadorRecibo
          End If
        End If
      Else 'Jurídica
        RECIBO.ReadOnly = True
        NFNUMERO.ReadOnly = False
        DATAEMISSAONOTA.ReadOnly = False

        If CurrentQuery.State <>1 Then
          CurrentQuery.FieldByName("RECIBO").Clear
        End If
      End If
    End If
    FinalizaSamParametrosProcContas
    Set qPrestador = Nothing
  End If
End Sub


Public Sub CHECATETOREEMBOLSO()
  'pronto deve checar o teto de reembolso pois se foi feito a revisão ele retira e não considera o teto
  If (CurrentQuery.FieldByName("SITUACAO").AsString = "3") And (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2) Then
    Dim spCalcTetoFinanciamento As BStoredProc
    Set spCalcTetoFinanciamento = NewStoredProc

    spCalcTetoFinanciamento.Name = "BSPROPEG_CALCTETOFINANCIAMENTO"
    spCalcTetoFinanciamento.AddParam("P_PEG",ptInput,ftInteger)
    spCalcTetoFinanciamento.AddParam("P_CHAVE",ptInput,ftInteger)
    spCalcTetoFinanciamento.AddParam("P_USUARIO",ptInput,ftInteger)
    spCalcTetoFinanciamento.ParamByName("P_PEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    spCalcTetoFinanciamento.ParamByName("P_CHAVE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    spCalcTetoFinanciamento.ParamByName("P_USUARIO").AsInteger = CurrentUser
    spCalcTetoFinanciamento.ExecProc

    Set spCalcTetoFinanciamento = Nothing
  End If
End Sub


Public Sub HabilitaEdicaoeBotoes
  Dim Especifico As Object
  Set Especifico = CreateBennerObject("Especifico.uEspecifico")
  If Especifico.Cliente(CurrentSystem) <> 5 Then 'FAPES
    If (CurrentQuery.FieldByName("SITUACAO").AsString = "3") And (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 2) Then
      Dim qVerificaValorPagtoConjuge As Object
      Set qVerificaValorPagtoConjuge = NewQuery

      qVerificaValorPagtoConjuge.Clear
      qVerificaValorPagtoConjuge.Add("SELECT GE.HANDLE ")
      qVerificaValorPagtoConjuge.Add("  FROM SAM_GUIA_EVENTOS GE")
      qVerificaValorPagtoConjuge.Add("  JOIN SAM_GUIA         GU ON GU.HANDLE = GE.GUIA")
      qVerificaValorPagtoConjuge.Add(" WHERE GU.PEG = :HANDLEPEG")
      qVerificaValorPagtoConjuge.Add("   AND VALORPAGTOCONJUGE > 0")
      qVerificaValorPagtoConjuge.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qVerificaValorPagtoConjuge.Active = True
      If qVerificaValorPagtoConjuge.FieldByName("HANDLE").AsInteger > 0 Then
        REVISAREVENTOS.Enabled = False
      End If
      Set qVerificaValorPagtoConjuge = Nothing
    End If
  End If
  Set Especifico = Nothing
End Sub


Public Function PegVinculadoImportacaoTISS(pClausulaWhere As String) As Boolean
  Dim qBuscaMensagemTiss As Object
  Set qBuscaMensagemTiss = NewQuery

  qBuscaMensagemTiss.Clear
  qBuscaMensagemTiss.Add("SELECT HANDLE FROM SAM_PRESTADOR_MENSAGEMTISS WHERE HANDLEPEG = :HANDLE " + pClausulaWhere)
  qBuscaMensagemTiss.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qBuscaMensagemTiss.Active = True
  If qBuscaMensagemTiss.FieldByName("HANDLE").AsInteger = 0 Then
    PegVinculadoImportacaoTISS = False
  Else
    PegVinculadoImportacaoTISS = True
  End If
  Set qBuscaMensagemTiss = Nothing
End Function


Public Sub CODIGOSERVICO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela  As String
  Dim vTitulo As String

  If CODIGOSERVICO.PopupCase <> 0 Then
    ShowPopup = False
    Set Interface = CreateBennerObject("Procura.Procurar")

    vCampos = "Código|Descrição Abreviada "
    vColunas = "CODIGO|DESCRICAOABREVIADA"
    vTabela = "SFN_CODSERVICOS"
    vTitulo = "Códigos de Serviços Município São Paulo"
    vCriterio = ""
    vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, vTitulo, True, CODIGOSERVICO.LocateText)

    Set Interface = Nothing
  Else
    ShowPopup = True
  End If

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CODIGOSERVICO").AsInteger = vHandle
    CurrentQuery.FieldByName("LISTASERVICO").Value = Null
  End If
End Sub


Public Sub LISTASERVICO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela  As String
  Dim vTitulo As String

  If LISTASERVICO.PopupCase <> 0 Then
    ShowPopup = False
    Set Interface = CreateBennerObject("Procura.Procurar")

    vCampos = "Código|Código Exportação|Descrição"
    vColunas = "CODIGO|CODIGOEXPORTACAO|DESCRICAOABREVIADA"
    vTabela = "SAM_LISTASERVICOS"
    vTitulo = "Lista de serviços"
    vCriterio = "(SAM_LISTASERVICOS.HANDLE IN (SELECT LISTASERVICO FROM SFN_CODSERVICOS_SERVICOSRELAC WHERE CODIGOSERVICO = " + CStr(CurrentQuery.FieldByName("CODIGOSERVICO").AsInteger) + "))"

    vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 3, vCampos, vCriterio, vTitulo, True, LISTASERVICO.LocateText)

    Set Interface = Nothing
  Else
    ShowPopup = True
  End If

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("LISTASERVICO").AsInteger = vHandle
  End If
End Sub


Public Function ValidaNumeroCartaRemessa() As Boolean
  If ((CurrentQuery.State = 2) Or (CurrentQuery.State = 3)) Then
    ValidaNumeroCartaRemessa = True
    'Verificação de Carta Remessa

    InicializaSamParametrosProcContas("REPETICAONUMERACAOPEG, SUGEREPEGCARTAREMESSA")
    If qSamParametrosProcContas.FieldByName("REPETICAONUMERACAOPEG").Value <> "1" _
       And qSamParametrosProcContas.FieldByName("SUGEREPEGCARTAREMESSA").AsBoolean = True Then

      If Not CurrentQuery.FieldByName("CARTAREMESSA").IsNull Then

        If CurrentQuery.FieldByName("SEQUENCIA").AsInteger =0 Then
          Dim CARTA As Object
          Set CARTA = NewQuery
          CARTA.Active = False
          CARTA.Clear
          CARTA.Add("SELECT HANDLE                       ")
          CARTA.Add("  FROM SAM_PEG                      ")
          CARTA.Add(" WHERE SEQUENCIA    =  0            ")
          CARTA.Add("   And CARTAREMESSA =  :CARTAREMESSA")
          CARTA.Add("   And SITUACAO	   <> '9'   	   ") 'sms 67652
          CARTA.Add("   And FILIAL       = :FILIAL")
          If CurrentQuery.State <> 3 Then
            CARTA.Add(" And HANDLE       <> :HANDLE      ")
          End If
          CARTA.ParamByName("FILIAL").AsInteger 		  = CurrentBranch  'CurrentQuery.FieldByName("FILIAL").AsFloat
          CARTA.ParamByName("CARTAREMESSA").AsFloat = CurrentQuery.FieldByName("CARTAREMESSA").AsFloat
          If CurrentQuery.State <> 3 Then
            CARTA.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
          End If
          CARTA.Active = True
          If Not CARTA.EOF Then
            bsShowMessage("Já existe um PEG com o mesmo número de Carta Remessa!", "E")
            ValidaNumeroCartaRemessa = False
          End If
          Set CARTA = Nothing
        End If
      End If
    End If
    FinalizaSamParametrosProcContas
  End If
End Function


Public Sub IncluiSessionVarMonitoramento()
  SessionVar("MONITORAMENTOPEG") = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("MONITORAMENTOGUIA") = "0"
End Sub


Public Function VerificaAgrupadorPagamentoFechado As Boolean
  If (CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
    Dim callEntity As CSEntityCall
    Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamPeg, Benner.Saude.Entidades", "VerificaPegVinculadoPagamentoFechado")
    callEntity.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("HANDLE").AsInteger)
    VerificaAgrupadorPagamentoFechado = CBool(callEntity.Execute)
    Set callEntity =  Nothing
  Else
    VerificaAgrupadorPagamentoFechado = False
  End If
End Function


Public Sub BOTAOENCAMINHAR_OnClick()
  Dim handlePeg As String
  handlePeg = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("HANDLE_PEG") = handlePeg

  If(CurrentQuery.FieldByName("TRIAGEMUSUARIO").AsInteger <> CurrentUser) And (Not CurrentQuery.FieldByName("TRIAGEMUSUARIO").IsNull) Then
  	bsShowMessage("Alteração não permitida para este usuário.", "I")
	Exit Sub
  End If

  If Not(IniciarValidacoesTriagem(handlePeg))Then
    Exit Sub
  End If

  If(VisibleMode) Then
    IniciarInterface
  End If
  AtualizarCarga(False)
End Sub


Public Sub IniciarInterface()
  Dim oForm As CSVirtualForm
  Set oForm = NewVirtualForm
  oForm.TableName = "TV_FORM0143"
  oForm.Caption   = "Encaminhamento de PEG´s"
  oForm.Height    = 230
  oForm.Width     = 300
  oForm.Show
  Set oForm = Nothing
End Sub


Public Function VerificarUsuarioParaEdicao As Boolean
  VerificarUsuarioParaEdicao = True
  If Not(VerificarUtilizaTriagem) Then
    Exit Function
  End If

  ' verifica o usuário é responsável pelo PEG
  If (VerificarMesmoUsuario) Then
    Exit Function
  End If

  ' verifica se o usuário é o digitador do PEG
  If (VerificarMesmoUsuarioDigitacao) Then
    Exit Function
  End If

  ' bloqueia a edição
  VerificarUsuarioParaEdicao = False
  bsShowMessage("Alteração não permitida para este usuário.", "I")
End Function


Public Function VerificarExisteUsuarioESetor As Boolean
  VerificarExisteUsuarioESetor = (Not(CurrentQuery.FieldByName("TRIAGEMSETOR").IsNull) And Not(CurrentQuery.FieldByName("TRIAGEMUSUARIO").IsNull))
End Function

Public Function VerificarMesmoUsuarioTriagem As Boolean
  VerificarMesmoUsuarioTriagem = (CurrentQuery.FieldByName("TRIAGEMUSUARIO").AsInteger = CurrentUser)
End Function


Public Function VerificarMesmoUsuarioDigitacao As Boolean
  VerificarMesmoUsuarioDigitacao = ((CurrentQuery.FieldByName("DIGITADOR").AsInteger = CurrentUser) And CurrentQuery.FieldByName("TRIAGEMUSUARIO").IsNull)
End Function


Public Sub HabilitarBotoesTriagem
  Dim bUtilizaTriagem As Boolean
  bUtilizaTriagem = VerificarUtilizaTriagem
  BOTAOENCAMINHAR.Enabled = bUtilizaTriagem
  BOTAOASSUMIRANALISE.Enabled = bUtilizaTriagem

  If Not(bUtilizaTriagem) Then
    Exit Sub
  End If

  bUtilizaTriagem = VerificarMesmoUsuario
  AtualizarBotoesProvisionamento(bUtilizaTriagem)

  If (bUtilizaTriagem) Then
  'se for para habilitar os botões, sai do processo para manter configurações existentes
    Exit Sub
  End If

  BOTAODIGITAR.Enabled = VerificarMesmoUsuarioDigitacao

  HabilitarDemais(bUtilizaTriagem)
  HabilitarFase(bUtilizaTriagem)
  HabilitarConferir(bUtilizaTriagem)
End Sub

Public Sub AtualizarBotoesProvisionamento(pbHabilita As Boolean)
  BOTAOCANCELARPROVISAO.Enabled = pbHabilita
  BOTAOVERIFICAMONITORAMENTO.Enabled = pbHabilita
  BOTAOPROVISIONARPEG.Enabled = pbHabilita
End Sub

Public Function VerificarMesmoUsuario As Boolean
  VerificarMesmoUsuario = False

  If VerificarExisteUsuarioESetor Then
    VerificarMesmoUsuario = VerificarMesmoUsuarioTriagem
  End If
End Function


Public Sub HabilitarDemais(pbHabilita As Boolean)
  CONCILIARNOTA.Enabled = pbHabilita
  CONFERIDO.Enabled = pbHabilita
  CRITICARDIGITACAO.Enabled = pbHabilita
  DESDOBRAR.Enabled = pbHabilita
  DEVOLVERPEG.Enabled = pbHabilita
  IMPORTARBENNER.Enabled = pbHabilita
  IMPORTARSUS.Enabled = pbHabilita
  REPROCESSARPEG.Enabled = pbHabilita
  REVISAREVENTOS.Enabled = pbHabilita
  CONTATERCEIRO.Enabled = pbHabilita
  BOTAOCANCELAFATURAMENTO.Enabled = pbHabilita
  VOLTARSITUACAO.Enabled = pbHabilita
  BOTAOLIBERARVERIFICACAO.Enabled = pbHabilita
  BOTAOALTERARDATACONTABIL.Enabled = pbHabilita
  BOTAOALTERARDATAPAGAMENTO.Enabled = pbHabilita
  BOTAOALTERARGUIASAPRESENTADAS.Enabled = pbHabilita
  BOTAOALTERARVALORAPRESENTADO.Enabled = pbHabilita
  BOTAOCANCELARPROVISAO.Enabled = pbHabilita
  BOTAOALTERARDADOSNF.Enabled = pbHabilita
  BOTAOVERIFICAMONITORAMENTO.Enabled = pbHabilita
  BOTAOPROVISIONARPEG.Enabled = pbHabilita
End Sub


Public Sub HabilitarFase(pbHabilita As Boolean)
  FASEPEG.Enabled = pbHabilita
  FASEPEGTODOS.Enabled = pbHabilita
End Sub


Public Sub HabilitarConferir(pbHabilita As Boolean)
  CONFPEG.Enabled = pbHabilita
  CONFGUIA.Enabled = pbHabilita
  CONFEVENTO.Enabled = pbHabilita
End Sub


Public Sub BOTAOASSUMIRANALISE_OnClick()
  IniciarAssumirAnalise
  AtualizarCarga(False)
End Sub


Public Sub IniciarAssumirAnalise()

	Dim vDllBSPro006 As Object

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamPeg")

	vDllBSPro006.IniciarAssumirAnalise(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

	Set vDllBSPro006 = Nothing

End Sub


Public Sub AtualizarCarga(pbMostraMensagem As Boolean)
  If VisibleMode Then
    RefreshNodesWithTableInterfacePEG("SAM_PEG")
  ElseIf pbMostraMensagem Then
    bsShowMessage("Usuário sem permissão nesta filial.", "E")
  End If
End Sub


Public Sub ReprocessarPEGWeb()
  Dim Obj As Object
  Dim viRetorno As Long
  Dim vsMensagemErro As String
  Dim vvContainer As CSDContainer
  Set vvContainer = NewContainer
  vvContainer.AddFields("HANDLE:INTEGER;")
  vvContainer.Insert
  vvContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

  Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
  viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "SAMPEG", _
                                     "REPROCESSARPEG", _
                                     "Reprocessamento de PEG - |" + _
                                     "PEG n. |" + CStr(CurrentQuery.FieldByName("PEG").AsInteger) + "|", _
                                     0, _
                                     "SAM_PEG", _
                                     "", _
                                     "", _
                                     "", _
                                     "", _
                                     True, _
                                     vsMensagemErro, _
                                     vvContainer)

  Set vvContainer = Nothing
  Set Obj = Nothing
  If viRetorno = 0 Then
    bsShowMessage("Reprocessamento enviado para execução no servidor!", "I")
  Else
    bsShowMessage("Erro ao enviar reprocessamento	 para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
  End If
End Sub

Public Sub ApagarBeneficiariosGuiaEventos()
  Dim qApagarBeneficiarioGuia As Object
  Set qApagarBeneficiarioGuia = NewQuery

  qApagarBeneficiarioGuia.Active = False
  qApagarBeneficiarioGuia.Clear
  qApagarBeneficiarioGuia.Add(" UPDATE SAM_GUIA                          ")
  qApagarBeneficiarioGuia.Add("      SET BENEFICIARIO       =  NULL,     ")
  qApagarBeneficiarioGuia.Add("          DVCARTAO           =  NULL,     ")
  qApagarBeneficiarioGuia.Add("          IDADEBENEFICIARIO  =  NULL      ")
  qApagarBeneficiarioGuia.Add(" WHERE PEG = :HANDLE                      ")
  qApagarBeneficiarioGuia.ParamByName("HANDLE").AsString = CurrentQuery.FieldByName("HANDLE").AsString
  qApagarBeneficiarioGuia.ExecSQL

  qApagarBeneficiarioGuia.Active = False
  qApagarBeneficiarioGuia.Clear
  qApagarBeneficiarioGuia.Add(" UPDATE SAM_GUIA_EVENTOS                  ")
  qApagarBeneficiarioGuia.Add("      SET BENEFICIARIO       =  NULL,     ")
  qApagarBeneficiarioGuia.Add("          DVCARTAO           =  NULL,     ")
  qApagarBeneficiarioGuia.Add("          IDADEBENEFICIARIO  =  NULL      ")
  qApagarBeneficiarioGuia.Add(" WHERE GUIA IN                            ")
  qApagarBeneficiarioGuia.Add("     (SELECT HANDLE                       ")
  qApagarBeneficiarioGuia.Add("         FROM SAM_GUIA                    ")
  qApagarBeneficiarioGuia.Add("      WHERE PEG = :HANDLE )               ")
  qApagarBeneficiarioGuia.Add("  AND ORIGEMEVENTO = :ORIGEMEVENTO        ")
  qApagarBeneficiarioGuia.ParamByName("HANDLE").AsString = CurrentQuery.FieldByName("HANDLE").AsString
  qApagarBeneficiarioGuia.ParamByName("ORIGEMEVENTO").AsString = "N"
  qApagarBeneficiarioGuia.ExecSQL

  Set qApagarBeneficiarioGuia = Nothing
End Sub


Public Sub InicializaSamParametrosProcContas(NomeDoCampo As String)
  'Utilizar esta Sub para acessar parametros gerais em seu método, atenção com saidas para outros métodos antes de finalizar.
  Set qSamParametrosProcContas = NewQuery
  qSamParametrosProcContas.Active = False
  qSamParametrosProcContas.Clear
  qSamParametrosProcContas.Add("  SELECT " + NomeDoCampo                   )
  qSamParametrosProcContas.Add("      FROM SAM_PARAMETROSPROCCONTAS       ")
  qSamParametrosProcContas.Active = True
End Sub


Public Sub FinalizaSamParametrosProcContas()
  qSamParametrosProcContas.Active = False
  Set qSamParametrosProcContas = Nothing

End Sub

Public Function ValidarEmpenho(handleEmpenho As Long) As Boolean

  ValidarEmpenho = True

  Dim qDadosEmpenho As BPesquisa
  Set qDadosEmpenho = NewQuery

  qDadosEmpenho.Clear
  qDadosEmpenho.Add("SELECT TABTIPO, PRESTADOR")
  qDadosEmpenho.Add("  FROM SFN_EMPENHO ")
  qDadosEmpenho.Add(" WHERE HANDLE = :HANDLE")
  qDadosEmpenho.ParamByName("HANDLE").AsInteger = handleEmpenho
  qDadosEmpenho.Active = True
  If (qDadosEmpenho.FieldByName("TABTIPO").AsInteger = 2) And (qDadosEmpenho.FieldByName("PRESTADOR").AsInteger > CurrentQuery.FieldByName("RECEBEDOR").AsInteger) Then
    ValidarEmpenho = False
  End If

  Set qDadosEmpenho = Nothing
End Function


Public Function ValidarEmpenhoConformeRecebedor() As Boolean

  ValidarEmpenhoConformeRecebedor = True

  If (CurrentQuery.FieldByName("TABREGIMEPGTO").AsInteger = 1) And (CurrentQuery.FieldByName("TABORIGEMRECURSOPEG").AsInteger = 2) Then ' Credenciamento e Origem Informado Orçamento

    ValidarEmpenhoConformeRecebedor = ValidarEmpenho(CurrentQuery.FieldByName("EMPENHOPEG").AsInteger)

  End If

End Function

Public Function VerificarRecebedorUtilizaAdiantamentoAutomatico() As Boolean
  Dim qPrestador As Object

  Set qPrestador = NewQuery
  qPrestador.Active = False
  qPrestador.Clear
  qPrestador.Add("  SELECT ADIANTAMENTOAUTOMATICO FROM SAM_PRESTADOR WHERE HANDLE = :HRECEBEDOR ")
  qPrestador.ParamByName("HRECEBEDOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  qPrestador.Active = True

  VerificarRecebedorUtilizaAdiantamentoAutomatico = qPrestador.FieldByName("ADIANTAMENTOAUTOMATICO").AsBoolean

  Set qPrestador = Nothing
End Function

Public Sub AtualizarDescontoINSS

	Dim qBuscaDesconto As Object

	Set qBuscaDesconto = NewQuery
	qBuscaDesconto.Active = False
	qBuscaDesconto.Clear

	qBuscaDesconto.Add("SELECT PERCENTUALREDUCAOINSS PERCENTUAL ")
	qBuscaDesconto.Add("  FROM SAM_PRESTADOR                    ")
	qBuscaDesconto.Add(" WHERE HANDLE = :RECEBEDOR              ")
	qBuscaDesconto.Add("   AND REDUCAOBASEINSS = 1              ")
	qBuscaDesconto.ParamByName("RECEBEDOR").AsInteger = CurrentQuery.FieldByName("RECEBEDOR").AsInteger

	qBuscaDesconto.Active = True

	If Not qBuscaDesconto.FieldByName("PERCENTUAL").IsNull Then
		CurrentQuery.FieldByName("PERCENTUALREDUCAOINSS").AsFloat = qBuscaDesconto.FieldByName("PERCENTUAL").AsFloat
	Else
		CurrentQuery.FieldByName("PERCENTUALREDUCAOINSS").Value = Null
	End If

	Set qBuscaDesconto = Nothing
End Sub

Private Sub RefreshPeg
  If VisibleMode Then
    RefreshNodesWithTable("SAM_PEG")
  End If
End Sub

Public Function EstaNaInterfaceDeDigitacao() As Boolean

	Dim vDllBSPro006 As Object

	Set vDllBSPro006 = CreateBennerObject("BSPro006.Rotinas")

  	EstaNaInterfaceDeDigitacao = vDllBSPro006.EstaNaInterfaceDeDigitacao(CurrentSystem)

  	Set vDllBSPro006 = Nothing

End Function

Public Function PermitirDigitacaoPeg() As Boolean

	Dim vDllBSPro006 As Object

	Set vDllBSPro006 = CreateBennerObject("BSPro006.Rotinas")

	PermitirDigitacaoPeg = vDllBSPro006.PermitirDigitacaoPeg(CurrentSystem)

	Set vDllBSPro006 = Nothing

End Function
