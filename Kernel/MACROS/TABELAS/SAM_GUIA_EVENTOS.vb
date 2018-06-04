'HASH: 70DE2DFB6AF47F99EB84690CCB72F6B6
'MACRO='Macro: SAM_GUIA_EVENTOS
Option Explicit

Dim vBeneficiario           As Long
Dim vValorApresentado       As Double
Dim vDvCartao               As String
Dim vExecutor               As Long
Dim vData                   As Date
Dim vQtd                    As Currency
Dim vEvento                 As Long
Dim vGrau                   As Long
Dim vCodigoPagto            As Long
Dim vPercentualdesconto     As Double
Dim interfacegeral          As Object
Dim OLDEVENTO               As Long
Dim OldSenha                As String
Dim gTipoAlteracao          As String
Dim old_modeloguia          As Long
Dim vSituacaoAnteriorGuia   As String
Dim vSituacaoAnteriorPeg    As String
Dim vSituacaoAnteriorEvento As String
Dim viState                 As Long
Dim vDllBSPro006            As Object

'#Uses "*ListaCamposLeiaute"
'#Uses "*ProcuraBeneficiarioAtivo"
'#Uses "*ProcuraBeneficiarioAtivoReembolso"
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"
'#Uses "*ProcuraPrestador"
'#Uses "*ProcuraGrauValido"
'#Uses "*ProcuraCodigoPagto"
'#Uses "*IsInt"
'#Uses "*PermissaoAlteracao"
'#Uses "*TV_FORM0143_VALIDACAO"
'#Uses "*RecordHandleOfTableInterfacePEG"

' Botões ------------------------------------------------------------

Public Sub BOTAOALERTA_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOALERTA_OnClick(CurrentSystem, CurrentQuery.TQuery, vSituacaoAnteriorEvento)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOALTERARVALORINFORMADOPF_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOALTERARVALORINFORMADOPF_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOAPAGAREVENTO_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOAPAGAREVENTO_OnClick(CurrentSystem, CurrentQuery.TQuery, vSituacaoAnteriorEvento)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOCANCELARPROVISAO_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOCANCELARPROVISAO_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOEVENTOORIGINAL_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOEVENTOORIGINAL_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAONOVOPERCENTUAL_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAONOVOPERCENTUAL_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOPFINTEGRAL_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOPFINTEGRAL_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOPRECOEVENTO_OnClick()

  	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOPRECOEVENTO_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOREGULARIZAR_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOREGULARIZAR_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub BOTAOREPROCESSAR_OnClick()

	BOTAOREPROCESSAR.Enabled = False

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.BOTAOREGULARIZAR_OnClick(CurrentSystem, CurrentQuery.TQuery)

	Set vDllBSPro006 = Nothing

	BOTAOREPROCESSAR.Enabled = True

End Sub

' Campos ------------------------------------------------------------

Public Sub TABTIPOGUIA_OnChanging(AllowChange As Boolean)
  MsgBox("Alteração não permitida")
  AllowChange = False
End Sub

Public Sub EXECUTOR_OnPopup(ShowPopup As Boolean)
  MeuBeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If
  '  If Len(EXECUTOR.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", EXECUTOR.Text)' pelo CPF e EXECUTOR
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EXECUTOR").Value = vHandle
  End If
  '  End If
End Sub

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vTexto As String
  Dim vRecebedor As Long
  Dim qRecebedor As Object
  Dim qParametros As Object
  Dim vUtilizarConsultaCentral As Boolean
  Dim vPermiteReembolsoFamiliar As Boolean 'Luciano T. Alberti - SMS 61716 - 03/05/2006
  Dim qDVCartao As Object 'SMS 68860 - Marcelo Barbosa - 06/10/2006

  vTexto = BENEFICIARIO.LocateText

  MeuBeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If
  '  If Len(BENEFICIARIO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  'vHandle = ProcuraBeneficiarioAtivo(False, ServerDate, vTexto)

  'Alterando na SMS 43212 - 07.03.2006
  Dim InterfaceBenef As Object
  Set InterfaceBenef = CreateBennerObject("CA010.ConsultaBeneficiario")
  If (Not CurrentQuery.FieldByName("RECEBEDOR").IsNull) Then
    vRecebedor = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
  Else
    Set qRecebedor = NewQuery
    qRecebedor.Add("SELECT P.RECEBEDOR RECEBEDORPEG, ")
    qRecebedor.Add("       G.RECEBEDOR RECEBEDORGUIA ")
    qRecebedor.Add("  FROM SAM_PEG P, ")
    qRecebedor.Add("       SAM_GUIA G ")
    qRecebedor.Add(" WHERE G.PEG = P.HANDLE ")
    qRecebedor.Add("   AND G.HANDLE = :GUIA ")
    qRecebedor.ParamByName("GUIA").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
    qRecebedor.Active = True
    If (Not qRecebedor.FieldByName("RECEBEDORGUIA").IsNull) Then
      vRecebedor = qRecebedor.FieldByName("RECEBEDORGUIA").AsInteger
    Else
      vRecebedor = qRecebedor.FieldByName("RECEBEDORPEG").AsInteger
    End If
    Set qRecebedor = Nothing
  End If

  Set qParametros = NewQuery
  qParametros.Add("SELECT UTILIZARCONSULTACENTRAL FROM SAM_PARAMETROSATENDIMENTO")
  qParametros.Active = True
  vUtilizarConsultaCentral = (qParametros.FieldByName("UTILIZARCONSULTACENTRAL").AsString = "S")

  'Luciano T. Alberti - SMS 61716 - 03/05/2006 - Início
  qParametros.Active = False
  qParametros.Clear
  qParametros.Add("Select PERMITEREEMBOLSOFAMILIAR FROM SAM_PARAMETROSPROCCONTAS")
  qParametros.Active = True
  vPermiteReembolsoFamiliar = (qParametros.FieldByName("PERMITEREEMBOLSOFAMILIAR").AsString = "S")
  'Luciano T. Alberti - SMS 61716 - 03/05/2006 - Fim

  Set qParametros = Nothing

  'SMS: 59436 - Rodrigo Soares - Início - Originava erro na pesquisa de beneficiario - Motivo: falta de parametro
  If (CurrentQuery.FieldByName("DATAATENDIMENTO").IsNull) Then
    InterfaceBenef.AlteraDataAtend(ServerDate, vRecebedor)
  Else
    InterfaceBenef.AlteraDataAtend(CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime, vRecebedor)
  End If
  'Rodrigo Soares - SMS 59436 - Fim

  'vHandle = InterfaceBenef.Filtro(CurrentSystem, 1, "")
  'Set InterfaceBenef = Nothing
  'Final SMS 43212

  'Dim vBenef As Integer 'Luciano T. Alberti - SMS 61716 - 03/05/2006
  Dim vBenef As Long 'Luciano T. Alberti - SMS 64243 - 28/06/2006

  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT P.TABREGIMEPGTO,")
  sql.Add("       P.BENEFICIARIO BenefPEG,")  'Luciano T. Alberti - SMS 61716 - 03/05/2006
  sql.Add("       G.BENEFICIARIO BenefGuia")  'Luciano T. Alberti - SMS 61716 - 03/05/2006
  sql.Add("  FROM SAM_PEG  P ")
  sql.Add("  JOIN SAM_GUIA G ON (G.PEG = P.HANDLE)")
  sql.Add(" WHERE G.HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
  sql.Active = True

  'Luciano T. Alberti - SMS 61716 - 03/05/2006 - Início
  If sql.FieldByName("BenefPEG").AsInteger > 0 Then
    vBenef = sql.FieldByName("BenefPEG").AsInteger
  ElseIf sql.FieldByName("BenefGUIA").AsInteger > 0 Then
    If vPermiteReembolsoFamiliar Then
      vBenef = sql.FieldByName("BenefGUIA").AsInteger
    Else
      vBenef = 0
    End If
  Else
    vBenef = 0
  End If
  'Luciano T. Alberti - SMS 61716 - 03/05/2006 - Fim

  If sql.FieldByName("TABREGIMEPGTO").AsInteger = 2 Then
    If vBenef > 0 Then 'Luciano T. Alberti - SMS 61716 - 03/05/2006
      If (vUtilizarConsultaCentral) Then
        vHandle = InterfaceBenef.FiltroTitular(CurrentSystem, 1, "", vBenef) 'Luciano T. Alberti - SMS 61716 - 03/05/2006
      Else
        vHandle = ProcuraBeneficiarioAtivoReembolso(False,ServerDate,BENEFICIARIO.LocateText, vBenef, True) 'Luciano T. Alberti - SMS 61716 - 03/05/2006
      End If

    Else
      'SMS 60629 - Marcelo Barbosa - 12/04/2006
      'If (vUtilizarConsultaCentral) Then
      '  vHandle = InterfaceBenef.Filtro(CurrentSystem, 1, "")
      'Else
        vHandle =ProcuraBeneficiarioAtivo(False,ServerDate,BENEFICIARIO.LocateText)
      'End If
      'Fim - SMS 60629
    End If
  Else
    'SMS 60629 - Marcelo Barbosa - 12/04/2006
    'If (vUtilizarConsultaCentral) Then
    '  vHandle = InterfaceBenef.Filtro(CurrentSystem, 1, "")
    'Else
      vHandle =ProcuraBeneficiarioAtivo(False,ServerDate,BENEFICIARIO.LocateText)
    'End If
    'Fim - SMS 60629
  End If

  Set InterfaceBenef = Nothing


  If (vHandle <> 0) Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle

    'SMS 68660 - Marcelo Barbosa - 06/10/2006
    Set qDVCartao = NewQuery

    If CurrentQuery.State <> 1 Then
      qDVCartao.Clear
      qDVCartao.Add("SELECT DVCARTAO FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEFICIARIO")
      qDVCartao.ParamByName("HBENEFICIARIO").AsInteger = vHandle
      qDVCartao.Active = True
      If Not qDVCartao.FieldByName("DVCARTAO").IsNull Then
        CurrentQuery.FieldByName("DVCARTAO").AsString = qDVCartao.FieldByName("DVCARTAO").AsString
      End If
    End If

    Set qDVCartao = Nothing
    'Fim - SMS 68660
  End If
  '  End If
End Sub

Public Sub CODIGOPAGTO_OnPopup(ShowPopup As Boolean)
  MeuBeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If
  '  If Len(CODIGOPAGTO.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraCodigoPagto(CODIGOPAGTO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CODIGOPAGTO").Value = vHandle
  End If
  '  End If
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  Dim vTexto As String

  vTexto = EVENTO.LocateText

  If ShowPopup = False Then
    Exit Sub
  End If

  Dim vHandle As Long
  ShowPopup = False

  If (CurrentQuery.FieldByName("CODIGOTABELA").AsInteger > 0) Then
  	vHandle = ProcuraEventoAtivoInativo(True, vTexto, CurrentQuery.FieldByName("CODIGOTABELA").AsInteger)
  Else
    vHandle = ProcuraEventoAtivoInativo(True, vTexto, 0)
  End If

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  MeuBeforeEdit(ShowPopup)
  If ShowPopup = False Then
    Exit Sub
  End If

  ShowPopup = False
  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    Exit Sub
  End If

  Dim vOrdemBusca As Integer '1 para buscar pelo grau e 2 para buscar pela descrição do grau
  Dim vbAux As Boolean

  Dim vBAchouGrauValido As Boolean
  vBAchouGrauValido = False

  vbAux = True
  On Error GoTo caracteres
  CDbl(GRAU.LocateText) 'sms 81953 - Artur - alterado de .Text para .LocateText
  vOrdemBusca = 1
  vbAux = False
	caracteres:
  If vbAux Then
    vOrdemBusca = 2
  End If

  Dim qParam As Object
  Set qParam = NewQuery

  On Error GoTo prox

  If GRAU.LocateText <> "" Then
    qParam.Clear
    qParam.Add("SELECT FILTRARGRAUSVALIDOSNADIGITACAO FROM SAM_PARAMETROSATENDIMENTO")
    qParam.Active = True
    If qParam.FieldByName("FILTRARGRAUSVALIDOSNADIGITACAO").AsString = "N" Then
      qParam.Clear

      'sms 81953 - Artur - após vericar sms 68642 ficou definido que o parâmetro "Filtrar Graus Válidos NA DIGITAÇÃO"
      'permite a seleção de qualquer grau APENAS NA DIGITAÇÃO, independente dos demais parâmentos, e não vale para a interface de Procura

      If vOrdemBusca = 1 Then
        'Luciano T. Alberti - SMS 68642 - 06/10/2006 - Início
        qParam.Add("SELECT G.HANDLE FROM SAM_GRAU G WHERE GRAU = :GRAU")
        qParam.ParamByName("GRAU").AsString = GRAU.LocateText
        'Luciano T. Alberti - SMS 68642 - 06/10/2006 - Fim
      Else
        'Luciano T. Alberti - SMS 68642 - 06/10/2006 - Início
        qParam.Add("SELECT G.HANDLE FROM SAM_GRAU G WHERE UPPER(G.DESCRICAO) LIKE UPPER('" + GRAU.LocateText + "%')")
        'Luciano T. Alberti - SMS 68642 - 06/10/2006 - Fim
      End If
      qParam.Active = True

      If Not qParam.EOF Then
        CurrentQuery.FieldByName("GRAU").AsInteger = qParam.FieldByName("HANDLE").AsInteger
        Set qParam = Nothing
        ShowPopup = False
        Exit Sub
      End If
    End If
  End If

	prox:
  If vBAchouGrauValido = True Then
    Exit Sub
  End If

  Dim vHandle As Long
  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim ProcuraGrauValido As Integer
  Dim TextoGrau As String
  TextoGrau = GRAU.LocateText 'sms 81953 - Artur - alterado de "" para GRAU.LocateText

  qParam.Clear
  qParam.Add("SELECT FILTRARGRAUSVALIDOS FROM SAM_PARAMETROSATENDIMENTO")
  qParam.Active = True

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "DISTINCT SAM_GRAU.GRAU|SAM_GRAU.CODIGOEXTERNO|SAM_GRAU.Z_DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"
  If qParam.FieldByName("FILTRARGRAUSVALIDOS").AsString = "S" Then
    vCriterio = "(SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM     SAM_TGE_GRAU GE WHERE GE.EVENTO=" + CurrentQuery.FieldByName("EVENTO").AsString + " AND     GE.GRAU=SAM_GRAU.HANDLE)))"
  Else
    vCriterio = ""
  End If
  vCampos = "Código do Grau|Código Externo|Descrição|Verifica grau válido"
  vTabela = "SAM_GRAU"

  ProcuraGrauValido = Interface.Exec(CurrentSystem, vTabela, vColunas, 3, vCampos, vCriterio, "Graus de Atuação", True, TextoGrau)

  Set Interface = Nothing
  'coelho Set qParam = Nothing


  If(ProcuraGrauValido <>0)Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = ProcuraGrauValido
    vBAchouGrauValido = True
    GRAU.LocateText = ""
  End If

	Set qParam = Nothing ' coelho

End Sub

Public Sub TIPOACOMODACAO_OnChange()
	  If (CurrentQuery.FieldByName("TIPOACOMODACAO").AsInteger > 0) And (ACOMODACAO.ReadOnly) Then
    Dim Q As Object
    Set Q = NewQuery
    Q.Add(" SELECT ACOMODACAO FROM TIS_TIPOACOMODACAO WHERE HANDLE = :HANDLE ")
    Q.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOACOMODACAO").AsInteger
    Q.Active = True
    If (Q.FieldByName("ACOMODACAO").AsInteger > 0) Then
      CurrentQuery.FieldByName("ACOMODACAO").AsInteger = Q.FieldByName("ACOMODACAO").AsInteger
    End If
    Set Q = Nothing
  End If
End Sub

Public Sub TIPOACOMODACAO_OnPopup(ShowPopup As Boolean)
    TIPOACOMODACAO.LocalWhere = " VERSAOTISS IN (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
End Sub

Public Sub VALORDESCONTO_OnExit()

	If (CurrentQuery.FieldByName("VALORDESCONTO").AsFloat < 0) Then
		bsShowMessage("O campo 'Vlr. Desc.' não pode ser negativo.", "I")
		VALORDESCONTO.SetFocus
	End If

End Sub

Public Sub BENEFICIARIO_OnExit()
  SugerirIdadeBeneficiario
End Sub

Public Sub CIDATENDIMENTO_OnEnter()
  CIDATENDIMENTO.AnyLevel = True
End Sub

Public Sub CIDATENDIMENTO_OnPopup(ShowPopup As Boolean)

  ' Leonardo Inicio 25/01/2001
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  ShowPopup = False
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CIDATENDIMENTO").Value = handlexx
  End If
  Set OLEAutorizador = Nothing
  ' Leonardo Fim 25/01/2001
End Sub

Public Sub CIDPRINCIPAL_OnEnter()
  CIDPRINCIPAL.AnyLevel = True
End Sub

Public Sub CIDPRINCIPAL_OnPopup(ShowPopup As Boolean)

  ' Leonardo Inicio 25/01/2001
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  ShowPopup = False
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  handlexx = OLEAutorizador.Exec(CurrentSystem, "SAM_CID", "ESTRUTURA|DESCRICAO", 2, "CID|Descrição", "HANDLE > 0", "Procura por CID", False, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CIDPRINCIPAL").Value = handlexx
  End If
  Set OLEAutorizador = Nothing
  ' Leonardo Fim 25/01/2001
End Sub

Public Sub CODIGOTABELA_OnPopup(ShowPopup As Boolean)
  	 CODIGOTABELA.LocalWhere = " VERSAOTISS IN (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
End Sub

Public Sub EVENTO_OnExit()
  If CurrentQuery.State <>1 Then
    Dim sql As Object
    Set sql = NewQuery
    sql.Add("SELECT COUNT(*) QTD FROM SAM_TGE_GRAU WHERE EVENTO = :HEVENTO")' And GRAUPRINCIPAL ='S' ")
    sql.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
    sql.Active = True
    If sql.FieldByName("QTD").AsInteger = 1 Then
      sql.Clear
      sql.Add("SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = :HEVENTO")' And GRAUPRINCIPAL ='S' ")
      sql.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
      sql.Active = True
      If Not sql.EOF Then
        CurrentQuery.FieldByName("GRAU").Value = sql.FieldByName("GRAU").AsInteger
      End If
    Else
      CurrentQuery.FieldByName("GRAU").Clear
    End If
    sql.Clear
    sql.Add("SELECT CODIGOPAGTO FROM SAM_TGE WHERE HANDLE = :HEVENTO")
    sql.ParamByName("HEVENTO").Value = CurrentQuery.FieldByName("EVENTO").AsInteger
    sql.Active = True
    If Not sql.FieldByName("CODIGOPAGTO").IsNull Then
      CurrentQuery.FieldByName("CODIGOPAGTO").AsInteger = sql.FieldByName("CODIGOPAGTO").AsInteger
    End If
    sql.Active = False
    Set sql = Nothing
  End If
'==============================================================================
'                       SMS 79066 - DRUMMOND - 26/04/2007 - INICIO
	Dim vSQL1 As Object

	Set vSQL1 = NewQuery
	vSQL1.Clear
	vSQL1.Add("SELECT CRITICARGRAUGUIAOUTRASDESP FROM SAM_PARAMETROSPROCCONTAS")
	vSQL1.Active = True

	If vSQL1.FieldByName("CRITICARGRAUGUIAOUTRASDESP").AsString = "S" Then
		'Verifico se o modelo da guia é do TISS...
		vSQL1.Active = False
		vSQL1.Clear
		vSQL1.Add("SELECT STG.TIPOGUIATISS                                          ")
		vSQL1.Add("  FROM SAM_GUIA              SG                                  ")
		vSQL1.Add("  JOIN SAM_TIPOGUIA_MDGUIA STMG ON (STMG.HANDLE = SG.MODELOGUIA) ")
		vSQL1.Add("  JOIN SAM_TIPOGUIA         STG ON (SG.HANDLE = STMG.TIPOGUIA)   ")
		vSQL1.Add(" WHERE SG.HANDLE = :HNDL                                         ")
		vSQL1.ParamByName("HNDL").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
		vSQL1.Active = True

		If vSQL1.FieldByName("TIPOGUIATISS").AsString <> "N" Then
			'Se a guia for diferente de "N" significa que ela pertence a um dos modelos do TISS
			vSQL1.Active = False
			vSQL1.Clear
            vSQL1.Add("SELECT EVENTODIARIA, GRAUDIARIA,           ")
            vSQL1.Add("       EVENTOTAXA, GRAUTAXA,               ")
            vSQL1.Add("       EVENTOMATERIAL, GRAUMATERIAL,       ")
            vSQL1.Add("       EVENTOMEDICAMENTO, GRAUMEDICAMENTO, ")
            vSQL1.Add("       EVENTOGASES, GRAUGASES,             ")
            vSQL1.Add("       EVENTOALUGUEIS, GRAUALUGUEIS        ")
            vSQL1.Add("  FROM TIS_PARAMETROS                      ")
			vSQL1.Active = True

			If (vSQL1.FieldByName("EVENTODIARIA").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger) And _
			(CurrentQuery.FieldByName("GRAU").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL1.FieldByName("GRAUDIARIA").AsInteger
			End If
			If (vSQL1.FieldByName("EVENTOTAXA").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger) And _
			(CurrentQuery.FieldByName("GRAU").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL1.FieldByName("GRAUTAXA").AsInteger
			End If
			If (vSQL1.FieldByName("EVENTOMATERIAL").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger) And _
			(CurrentQuery.FieldByName("GRAU").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL1.FieldByName("GRAUMATERIAL").AsInteger
			End If
			If (vSQL1.FieldByName("EVENTOMEDICAMENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger) And _
			(CurrentQuery.FieldByName("GRAU").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL1.FieldByName("GRAUMEDICAMENTO").AsInteger
			End If
			If (vSQL1.FieldByName("EVENTOGASES").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger) And _
			(CurrentQuery.FieldByName("GRAU").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL1.FieldByName("GRAUGASES").AsInteger
			End If
			If (vSQL1.FieldByName("EVENTOALUGUEIS").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger) And _
			(CurrentQuery.FieldByName("GRAU").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL1.FieldByName("GRAUALUGUEIS").AsInteger
			End If
		End If
	End If
	Set vSQL1 = Nothing
'                       SMS 79066 - DRUMMOND - 26/04/2007 - FIM
'==============================================================================
End Sub

Public Sub GRAU_OnExit()

	Dim vSQL1 As Object

	Set vSQL1 = NewQuery
	vSQL1.Clear
	vSQL1.Add("SELECT CRITICARGRAUGUIAOUTRASDESP FROM SAM_PARAMETROSPROCCONTAS")
	vSQL1.Active = True

	If vSQL1.FieldByName("CRITICARGRAUGUIAOUTRASDESP").AsString = "S" Then
		'Verifico se o modelo da guia é de Despesas...
		vSQL1.Active = False
		vSQL1.Clear
		vSQL1.Add("SELECT STG.TIPOGUIATISS                                          ")
		vSQL1.Add("  FROM SAM_GUIA              SG                                  ")
		vSQL1.Add("  JOIN SAM_TIPOGUIA_MDGUIA STMG ON (STMG.HANDLE = SG.MODELOGUIA) ")
		vSQL1.Add("  JOIN SAM_TIPOGUIA         STG ON (SG.HANDLE = STMG.TIPOGUIA)   ")
		vSQL1.Add(" WHERE SG.HANDLE = :HNDL                                         ")
		vSQL1.ParamByName("HNDL").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
		vSQL1.Active = True

		If vSQL1.FieldByName("TIPOGUIATISS").AsString <> "N" Then
			'Se a guia for diferente de "N" significa que ela pertence a um dos modelos do TISS
			vSQL1.Active = False
			vSQL1.Clear
            vSQL1.Add("SELECT EVENTODIARIA, GRAUDIARIA,           ")
            vSQL1.Add("       EVENTOTAXA, GRAUTAXA,               ")
            vSQL1.Add("       EVENTOMATERIAL, GRAUMATERIAL,       ")
            vSQL1.Add("       EVENTOMEDICAMENTO, GRAUMEDICAMENTO, ")
            vSQL1.Add("       EVENTOGASES, GRAUGASES,             ")
            vSQL1.Add("       EVENTOALUGUEIS, GRAUALUGUEIS        ")
            vSQL1.Add("  FROM TIS_PARAMETROS                      ")
			vSQL1.Active = True

			If (vSQL1.FieldByName("GRAUDIARIA").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger) And _
			(CurrentQuery.FieldByName("EVENTO").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("EVENTO").AsInteger = vSQL1.FieldByName("EVENTODIARIA").AsInteger
			End If
			If (vSQL1.FieldByName("GRAUTAXA").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger) And _
			(CurrentQuery.FieldByName("EVENTO").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("EVENTO").AsInteger = vSQL1.FieldByName("EVENTOTAXA").AsInteger
			End If
			If (vSQL1.FieldByName("GRAUMATERIAL").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger) And _
			(CurrentQuery.FieldByName("EVENTO").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("EVENTO").AsInteger = vSQL1.FieldByName("EVENTOMATERIAL").AsInteger
			End If
			If (vSQL1.FieldByName("GRAUMEDICAMENTO").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger) And _
			(CurrentQuery.FieldByName("EVENTO").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("EVENTO").AsInteger = vSQL1.FieldByName("EVENTOMEDICAMENTO").AsInteger
			End If
			If (vSQL1.FieldByName("GRAUGASES").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger) And _
			(CurrentQuery.FieldByName("EVENTO").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("EVENTO").AsInteger = vSQL1.FieldByName("EVENTOGASES").AsInteger
			End If
			If (vSQL1.FieldByName("GRAUALUGUEIS").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger) And _
			(CurrentQuery.FieldByName("EVENTO").AsInteger = 0) Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("EVENTO").AsInteger = vSQL1.FieldByName("EVENTOALUGUEIS").AsInteger
			End If
		End If
	End If

	If CurrentQuery.State <>1 Then
	  If (CurrentQuery.FieldByName("GRAU").AsInteger > 0) Then
        vSQL1.Clear
   	    vSQL1.Add("SELECT HANDLE                                    ")
   	    vSQL1.Add("  FROM TIS_DENTEFACE                             ")
   	    vSQL1.Add(" WHERE GRAU = :GRAU                              ")
   	    vSQL1.Add("   AND VERSAOTISS IN (SELECT MAX (HANDLE)        ")
        vSQL1.Add("                        FROM TIS_VERSAO          ")
        vSQL1.Add("                       WHERE ATIVODESKTOP = 'S') ")

        vSQL1.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
        vSQL1.Active = True
        If (vSQL1.FieldByName("HANDLE").AsInteger > 0) Then
          CurrentQuery.FieldByName("DENTEREGIAOFACE").AsInteger = vSQL1.FieldByName("HANDLE").AsInteger
	    Else
	  	  CurrentQuery.FieldByName("DENTEREGIAOFACE").Clear
        End If
      End If
	End If

	Set vSQL1 = Nothing
'                       SMS 79066 - DRUMMOND - 26/04/2007 - FIM
End Sub

Public Sub GRAUPARTICIPACAO_OnChange()
	Dim vSQL As Object

	Set vSQL = NewQuery
	vSQL.Clear
	vSQL.Add("SELECT UTILIZARGRAUPARTICIPACAO FROM SAM_PARAMETROSATENDIMENTO")
	vSQL.Active = True

	If (vSQL.FieldByName("UTILIZARGRAUPARTICIPACAO").AsString = "S") Then 'And (GRAUPARTICIPACAO.Text <> "") Then
		vSQL.Active = False
		vSQL.Clear
		vSQL.Add("SELECT GRAU FROM TIS_POSICAOPROFISSIONAL WHERE HANDLE = :HNDL ")
		vSQL.Add("                                           And VERSAOTISS In (Select MAX (HANDLE) ")
        vSQL.Add("                                                                FROM TIS_VERSAO ")
		vSQL.Add("                                                               WHERE ATIVODESKTOP = 'S') ")

		vSQL.ParamByName("HNDL").AsInteger = CurrentQuery.FieldByName("GRAUPARTICIPACAO").AsInteger 'CInt(GRAUPARTICIPACAO.Text)
		vSQL.Active = True

		If (vSQL.FieldByName("GRAU").AsInteger > 0) And (CurrentQuery.FieldByName("GRAU").AsInteger <= 0) Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("GRAU").AsInteger = vSQL.FieldByName("GRAU").AsInteger
		End If
	End If
	Set vSQL = Nothing
End Sub

Public Sub GRAUPARTICIPACAO_OnPopup(ShowPopup As Boolean)
	Dim viHandle As Long
	Dim Procura As Object
	Dim vSQL As Object
	Dim vbAux As Boolean
    Dim vOrdemBusca As Integer

	ShowPopup = False

    'SMS 84495 - Débora Rebello (CASSI) - 12/07/2007 - inicio
    vbAux = True
    On Error GoTo caracteres
      CDbl(GRAUPARTICIPACAO.LocateText)
      vOrdemBusca = 1
      vbAux = False
    caracteres:
      If vbAux Then
        vOrdemBusca = 2
      End If
    'SMS 84495 - Débora Rebello (CASSI) - 12/07/2007 - fim

    Dim vSelecaoEspecial As String		'SMS 104421 - 17/11/2008 - EVANDRO ZEFERINO
    vSelecaoEspecial = " TIS_POSICAOPROFISSIONAL.VERSAOTISS IN (SELECT PEG.VERSAOTISS FROM SAM_PEG PEG JOIN SAM_GUIA GUIA ON (PEG.HANDLE = GUIA.PEG) WHERE GUIA.HANDLE = " + CurrentQuery.FieldByName("GUIA").AsString + ")"

	Set Procura = CreateBennerObject("Procura.Procurar")
	viHandle = Procura.Exec(CurrentSystem, "TIS_POSICAOPROFISSIONAL", "CODIGO|DESCRICAO", vOrdemBusca, "Código|Descrição", vSelecaoEspecial, "Grau de participação", False, GRAUPARTICIPACAO.LocateText) 'SMS 84495 - Débora Rebello (CASSI) - 12/07/2007
	If viHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAUPARTICIPACAO").AsInteger = viHandle

		Set vSQL = NewQuery
		vSQL.Clear
		vSQL.Add("SELECT UTILIZARGRAUPARTICIPACAO FROM SAM_PARAMETROSATENDIMENTO")
		vSQL.Active = True

		If vSQL.FieldByName("UTILIZARGRAUPARTICIPACAO").AsString = "S" Then
			vSQL.Active = False
			vSQL.Clear
			vSQL.Add("SELECT GRAU FROM TIS_POSICAOPROFISSIONAL WHERE HANDLE = :HNDL")
			vSQL.Add("                                           And VERSAOTISS In (Select MAX (HANDLE) ")
    	    vSQL.Add("                                                                FROM TIS_VERSAO ")
			vSQL.Add("                                                               WHERE ATIVODESKTOP = 'S') ")
			vSQL.ParamByName("HNDL").AsInteger = viHandle
			vSQL.Active = True

			If (vSQL.FieldByName("GRAU").AsInteger > 0) And (CurrentQuery.FieldByName("GRAU").AsInteger <= 0)Then
				CurrentQuery.Edit
				CurrentQuery.FieldByName("GRAU").AsInteger = vSQL.FieldByName("GRAU").AsInteger
			End If
		End If
		Set vSQL = Nothing
	End If
End Sub

Public Sub RECEBEDOR_OnPopup(ShowPopup As Boolean)
  '#Uses "*ProcuraPrestador"
  '  If Len(RECEBEDOR.Text)=0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraPrestador("C", "T", RECEBEDOR.Text)' pelo CPF e RECEBEDOR
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("RECEBEDOR").Value = vHandle
  End If
End Sub

Public Sub SENHA_OnExit()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.SENHA_OnExit(CurrentSystem, CurrentQuery.TQuery, OldSenha)

	Set vDllBSPro006 = Nothing

End Sub

' Tabela -----------------------------------------------

Public Sub TABLE_AfterInsert()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_AfterInsert(CurrentSystem, CurrentQuery.TQuery, CurrentQuery.FieldByName("GUIA").AsInteger)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterPost()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_AfterPost(CurrentSystem, CurrentQuery.TQuery, viState, vSituacaoAnteriorPeg, vSituacaoAnteriorGuia, vSituacaoAnteriorEvento)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterScroll()

	MostraCamposLeiaute

	Dim vsEventoWebLocalWhere             As String
    Dim vsGrauWebLocalWhere               As String
    Dim vsCbosLocalWhere                  As String
    Dim vsCbosWebLocalWhere               As String
    Dim vsRotuloLimContagens              As String
    Dim vbFatorReducAcresAprReadOnly      As Boolean
    Dim vbTecnicaUtilizadaInformaReadOnly As Boolean
    Dim vbViaDeAcessoInformadaReadOnly    As Boolean
    Dim vbRecebedorVisible                As Boolean
    Dim vbBotaoRegularizarEnable          As Boolean
    Dim vbBotaoRegularizarVisible         As Boolean
    Dim vbBotaoReprocessarEnable          As Boolean
    Dim vbBotaoReprocessarVisible         As Boolean
    Dim vbBotaoNovoPercentualVisible      As Boolean
    Dim vbBotaoAlertaEnable               As Boolean
    Dim vbBotaoGerarGlosaEnable           As Boolean
    Dim vbBotaoIncluirPrestadorEnable     As Boolean
    Dim vbBotaoMatMedEnable               As Boolean
    Dim vbBotaoPfIntegralEnable           As Boolean
    Dim vbBotaoPrecoEventoEnable          As Boolean
    Dim vbBotaoAlterarValorInfPfVisible   As Boolean
    Dim vbBotaoCancelarProvisaoEnable     As Boolean
	Dim vbBotaoApagarEventoEnable         As Boolean
	Dim vbBotaoEventoOriginalEnable       As Boolean
    Dim vbTableReadOnly                   As Boolean

	vsEventoWebLocalWhere             = EVENTO.WebLocalWhere
	vsCbosLocalWhere                  = CBOS.LocalWhere
    vsCbosWebLocalWhere               = CBOS.WebLocalWhere
    vsGrauWebLocalWhere               = GRAU.WebLocalWhere
    vsRotuloLimContagens              = ROTULOLIMCONTAGENS.Text
    vbFatorReducAcresAprReadOnly      = FATORREDUCACRESCAPRESENTADO.ReadOnly
    vbTecnicaUtilizadaInformaReadOnly = TECNICAUTILIZADAINFORMADA.ReadOnly
    vbViaDeAcessoInformadaReadOnly    = VIADEACESSOINFORMADA.ReadOnly
    vbRecebedorVisible                = RECEBEDOR.Visible
    vbBotaoRegularizarEnable          = BOTAOREGULARIZAR.Enabled
    vbBotaoRegularizarVisible         = BOTAOREGULARIZAR.Visible
    vbBotaoReprocessarEnable          = BOTAOREPROCESSAR.Enabled
    vbBotaoReprocessarVisible         = BOTAOREPROCESSAR.Visible
    vbBotaoNovoPercentualVisible      = BOTAONOVOPERCENTUAL.Visible
    vbBotaoAlertaEnable               = BOTAOALERTA.Enabled
    vbBotaoGerarGlosaEnable           = BOTAOGERARGLOSA.Enabled
    vbBotaoIncluirPrestadorEnable     = BOTAOINCLUIRPRESTADOR.Enabled
    vbBotaoMatMedEnable               = BOTAOMATMED.Enabled
    vbBotaoPfIntegralEnable           = BOTAOPFINTEGRAL.Enabled
    vbBotaoPrecoEventoEnable          = BOTAOPRECOEVENTO.Enabled
    vbBotaoAlterarValorInfPfVisible   = BOTAOALTERARVALORINFORMADOPF.Visible
    vbBotaoCancelarProvisaoEnable     = BOTAOCANCELARPROVISAO.Visible
	vbBotaoApagarEventoEnable         = BOTAOAPAGAREVENTO.Enabled
	vbBotaoEventoOriginalEnable       = BOTAOEVENTOORIGINAL.Enabled
    vbTableReadOnly                   = TableReadOnly

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_AfterScroll(CurrentSystem, _
	                               CurrentQuery.TQuery, _
								   vsEventoWebLocalWhere, _
								   vsGrauWebLocalWhere, _
								   vsCbosLocalWhere, _
								   vsCbosWebLocalWhere, _
								   vsRotuloLimContagens, _
								   vbFatorReducAcresAprReadOnly, _
								   vbTecnicaUtilizadaInformaReadOnly, _
								   vbViaDeAcessoInformadaReadOnly, _
								   vbRecebedorVisible, _
								   vbBotaoRegularizarEnable, _
								   vbBotaoRegularizarVisible, _
								   vbBotaoReprocessarEnable, _
								   vbBotaoReprocessarVisible, _
								   vbBotaoNovoPercentualVisible, _
								   vbBotaoAlertaEnable, _
								   vbBotaoGerarGlosaEnable, _
								   vbBotaoIncluirPrestadorEnable, _
								   vbBotaoMatMedEnable, _
								   vbBotaoPfIntegralEnable, _
								   vbBotaoPrecoEventoEnable, _
								   vbBotaoAlterarValorInfPfVisible, _
								   vbBotaoCancelarProvisaoEnable, _
								   vbBotaoApagarEventoEnable, _
								   vbBotaoEventoOriginalEnable, _
								   vSituacaoAnteriorPeg, _
								   vSituacaoAnteriorGuia, _
								   vSituacaoAnteriorEvento, _
								   vbTableReadOnly)

	Set vDllBSPro006 = Nothing

	EVENTO.WebLocalWhere                 = vsEventoWebLocalWhere
    GRAU.WebLocalWhere                   = vsGrauWebLocalWhere
    CBOS.LocalWhere                      = vsCbosLocalWhere
    CBOS.WebLocalWhere                   = vsCbosWebLocalWhere
    ROTULOLIMCONTAGENS.Text              = vsRotuloLimContagens
    FATORREDUCACRESCAPRESENTADO.ReadOnly = vbFatorReducAcresAprReadOnly
    TECNICAUTILIZADAINFORMADA.ReadOnly   = vbTecnicaUtilizadaInformaReadOnly
    VIADEACESSOINFORMADA.ReadOnly        = vbViaDeAcessoInformadaReadOnly
    RECEBEDOR.Visible                    = vbRecebedorVisible
    BOTAOREGULARIZAR.Enabled             = vbBotaoRegularizarEnable
    BOTAOREGULARIZAR.Visible             = vbBotaoRegularizarVisible
    BOTAOREPROCESSAR.Enabled             = vbBotaoReprocessarEnable
    BOTAOREPROCESSAR.Visible             = vbBotaoReprocessarVisible
    BOTAONOVOPERCENTUAL.Visible          = vbBotaoNovoPercentualVisible
    BOTAOALERTA.Enabled                  = vbBotaoAlertaEnable
    BOTAOGERARGLOSA.Enabled              = vbBotaoGerarGlosaEnable
    BOTAOINCLUIRPRESTADOR.Enabled        = vbBotaoIncluirPrestadorEnable
    BOTAOMATMED.Enabled                  = vbBotaoMatMedEnable
    BOTAOPFINTEGRAL.Enabled              = vbBotaoPfIntegralEnable
    BOTAOPRECOEVENTO.Enabled             = vbBotaoPrecoEventoEnable
    BOTAOALTERARVALORINFORMADOPF.Visible = vbBotaoAlterarValorInfPfVisible
    BOTAOCANCELARPROVISAO.Visible        = vbBotaoCancelarProvisaoEnable
	BOTAOAPAGAREVENTO.Enabled            = vbBotaoApagarEventoEnable
	BOTAOEVENTOORIGINAL.Enabled          = vbBotaoEventoOriginalEnable
    TableReadOnly                        = vbTableReadOnly

    MOTIVOREAPRESENTACAO.ReadOnly        = CurrentQuery.FieldByName("EVENTOORIGINAL").IsNull
    TEXTOMOTIVOREAPRESENTACAO.ReadOnly   = CurrentQuery.FieldByName("EVENTOORIGINAL").IsNull
	GLOSATISSREAPRESENTACAO.ReadOnly     = CurrentQuery.FieldByName("EVENTOORIGINAL").IsNull

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_BeforeDelete(CurrentSystem, CurrentQuery.TQuery, vSituacaoAnteriorEvento, CanContinue)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

	Dim vbBotaoRegularizarEnabled   As Boolean
	Dim vbBotaoReprocessarEnabled   As Boolean
	Dim vbBotaoAlertaEnabled        As Boolean

	vbBotaoRegularizarEnabled   = BOTAOREGULARIZAR.Enabled
	vbBotaoReprocessarEnabled   = BOTAOREPROCESSAR.Enabled
	vbBotaoAlertaEnabled        = BOTAOALERTA.Enabled

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_BeforeEdit(CurrentSystem, _
	                              CurrentQuery.TQuery, _
	                              CanContinue, _
								  vBeneficiario, _
								  vValorApresentado, _
								  vDvCartao, _
								  vExecutor, _
								  vData, _
								  vQtd, _
								  vEvento, _
								  vGrau, _
								  vCodigoPagto, _
								  OLDEVENTO, _
                                  OldSenha, _
								  vPercentualdesconto, _
								  vbBotaoRegularizarEnabled, _
								  vbBotaoReprocessarEnabled, _
								  vbBotaoAlertaEnabled)
	Set vDllBSPro006 = Nothing

	BOTAOREGULARIZAR.Enabled = vbBotaoRegularizarEnabled
	BOTAOREPROCESSAR.Enabled = vbBotaoReprocessarEnabled
	BOTAOALERTA.Enabled      = vbBotaoAlertaEnabled
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_BeforeInsert(CurrentSystem, CurrentQuery.TQuery, RecordHandleOfTableInterfacePEG("SAM_PEG"), CanContinue)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_BeforePost(CurrentSystem, _
	                              CurrentQuery.TQuery, _
	                              vBeneficiario, _
							      vValorApresentado, _
								  vExecutor, _
								  vData, _
								  vQtd, _
								  vCodigoPagto, _
								  vEvento, _
								  vPercentualdesconto, _
								  vGrau, _
								  CanContinue, _
								  viState, _
								  gTipoAlteracao, _
								  OldSenha)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_AfterDelete()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_AfterDelete(CurrentSystem, _
	                               CurrentQuery.TQuery, _
								   vSituacaoAnteriorGuia, _
		                           vSituacaoAnteriorPeg, _
		                           vSituacaoAnteriorEvento)
	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_Aftercommitted()

  	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_Aftercommitted(CurrentSystem, _
	                                  CurrentQuery.TQuery, _
								      gTipoAlteracao)
	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_BeforeScroll()

	VALORPAGOWEB.Visible = Not WebMode

End Sub

Public Sub TABLE_NewRecord()

	CODIGOTABELA.LocalWhere   = " VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S')"
	TIPOACOMODACAO.LocalWhere = " VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S') "

  	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.TABLE_NewRecord(CurrentSystem, _
	                             CurrentQuery.TQuery, _
								 vSituacaoAnteriorPeg, _
								 vSituacaoAnteriorGuia, _
								 vSituacaoAnteriorEvento)
	Set vDllBSPro006 = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	Select Case CommandID
		Case "BOTAOREPROCESSAR"
			BOTAOREPROCESSAR_OnClick
	 	Case "BOTAOREGULARIZAR"
	 		BOTAOREGULARIZAR_OnClick
	 	Case "BOTAOPFINTEGRAL"
	 	    BOTAOPFINTEGRAL_OnClick
	 	Case "BOTAOAPAGAREVENTO"
	 		BOTAOAPAGAREVENTO_OnClick
	 	Case "BOTAOPRECOEVENTO"
	 	    BOTAOPRECOEVENTO_OnClick
	End Select
End Sub

' Funções --------------------------------------

Public Sub EscondeCamposLeiaute

  'esses são os campos que estariam no panel da SamPegDigit.dll que é montado conforme o modelo de GUIA
  DVCARTAO.ReadOnly = True
  PARCELAMENTO.ReadOnly = True
  HORARIOESPECIAL.ReadOnly = True
  PERCENTUALDESCONTO.ReadOnly = True
  ORDEM.ReadOnly = True
  BENEFICIARIO.ReadOnly = True
  EVENTO.ReadOnly = True
  EXECUTOR.ReadOnly = True
  GRAU.ReadOnly = True
  CODIGOPAGTO.ReadOnly = True
  XTHM.ReadOnly = True
  CIDPRINCIPAL.ReadOnly = True
  CIDATENDIMENTO.ReadOnly = True
  FINALIDADEATENDIMENTO.ReadOnly = True
  RECEBEDOR.ReadOnly = True
  DATAATENDIMENTO.ReadOnly = True
  HORAATENDIMENTO.ReadOnly = Not CurrentQuery.InInsertion
  IDADEBENEFICIARIO.ReadOnly = True
  CHECONSULTA.ReadOnly = True
  ACOMODACAO.ReadOnly = True
  TIPOACOMODACAO.ReadOnly = True  'SMS 77448 - 24.04.2007

End Sub

Public Sub MostraCamposLeiaute
  If CurrentQuery.FieldByName("SITUACAO").AsString = "1" Or CurrentQuery.FieldByName("SITUACAO").AsString = "4" Then
    Exit Sub
  End If

  Dim Q1 As Object
  Dim Modelo As Long
  Dim vAux As String
  Set Q1 = NewQuery
  Q1.Add("SELECT MODELOGUIA FROM SAM_GUIA WHERE HANDLE=" + Str(RecordHandleOfTableInterfacePEG("SAM_GUIA")))
  Q1.Active = True

  Modelo = Q1.FieldByName("MODELOGUIA").AsInteger
  Q1.Active = False
  Set Q1 = Nothing

  If old_modeloguia = Modelo Then
    Exit Sub
  End If


  ListaCamposLeiaute Modelo, 2

  EscondeCamposLeiaute

  vAux = UserVar("CAMPOS_LEIAUTE_GUIA_EVENTO")
  ' MsgBox(vAux)


  If(InStr(vAux, "|" + "DVCARTAO")>0)And(CurrentQuery.RequestLive)Then DVCARTAO.ReadOnly = False
  If(InStr(vAux, "|" + "PARCELAMENTO")>0)And(CurrentQuery.RequestLive)Then PARCELAMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "HORARIOESPECIAL")>0)And(CurrentQuery.RequestLive)Then HORARIOESPECIAL.ReadOnly = False
  If(InStr(vAux, "|" + "PERCENTUALDESCONTO")>0)And(CurrentQuery.RequestLive)Then PERCENTUALDESCONTO.ReadOnly = False
  If(InStr(vAux, "|" + "ORDEM")>0)And(CurrentQuery.RequestLive)Then ORDEM.ReadOnly = False
  If(InStr(vAux, "|" + "BENEFICIARIO")>0)And(CurrentQuery.RequestLive)Then BENEFICIARIO.ReadOnly = False
  If(InStr(vAux, "|" + "EVENTO")>0)And(CurrentQuery.RequestLive)Then EVENTO.ReadOnly = False
  If(InStr(vAux, "|" + "EXECUTOR")>0)And(CurrentQuery.RequestLive)Then EXECUTOR.ReadOnly = False
  If(InStr(vAux, "|" + "GRAU")>0)And(CurrentQuery.RequestLive)Then GRAU.ReadOnly = False
  If(CurrentQuery.RequestLive)Then CODIGOPAGTO.ReadOnly = False
  If(CurrentQuery.RequestLive)Then XTHM.ReadOnly = False
  If(InStr(vAux, "|" + "CIDPRINCIPAL")>0)And(CurrentQuery.RequestLive)Then CIDPRINCIPAL.ReadOnly = False
  If(InStr(vAux, "|" + "CIDATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then CIDATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "FINALIDADEATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then FINALIDADEATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "RECEBEDOR")>0)And(CurrentQuery.RequestLive)Then RECEBEDOR.ReadOnly = False
  If(InStr(vAux, "|" + "DATAATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then DATAATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "HORAATENDIMENTO")>0)And(CurrentQuery.RequestLive)Then HORAATENDIMENTO.ReadOnly = False
  If(InStr(vAux, "|" + "IDADEBENEFICIARIO")>0)And(CurrentQuery.RequestLive)Then IDADEBENEFICIARIO.ReadOnly = False
  If(InStr(vAux, "|" + "CHECONSULTA")>0)And(CurrentQuery.RequestLive)Then CHECONSULTA.ReadOnly = False
  If(InStr(vAux, "|" + "ACOMODACAO")>0)And(CurrentQuery.RequestLive)Then ACOMODACAO.ReadOnly = False
  If(InStr(vAux, "|" + "TIPOACOMODACAO")>0)And(CurrentQuery.RequestLive)Then TIPOACOMODACAO.ReadOnly = False  'SMS 77448 - 24.04.2007



End Sub

Function CalculaIdadeBeneficiario(ByVal pBeneficiario As Long, ByVal pDataAtendimento As Date)As Integer
  Dim vDias           As Integer
  Dim vMeses          As Integer
  Dim vAnos           As Integer
  Dim VDataNascimento As Date
  Dim query As Object


  Set query = NewQuery
  query.Clear
  query.Add("SELECT M.DATANASCIMENTO ")
  query.Add("  FROM SAM_MATRICULA M ")
  query.Add("  JOIN SAM_BENEFICIARIO B ON B.MATRICULA = M.HANDLE ")
  query.Add(" WHERE B.HANDLE  = :HANDLE ")
  query.ParamByName("HANDLE").AsInteger = pBeneficiario
  query.Active = True

  VDataNascimento = query.FieldByName("DATANASCIMENTO").AsDateTime
  If (VDataNascimento > ServerDate) Then
    CalculaIdadeBeneficiario = 0
  Else
    DiferencaData2 pDataAtendimento, VDataNascimento, vDias, vMeses, vAnos
    CalculaIdadeBeneficiario = vAnos
  End If
End Function

Public Sub DiferencaData2(ByVal Data1, Data2 As Date, Dias, Meses, Anos As Integer)
  Dim DtSwap As Date
  Dim Day1, Day2, Month1, Month2, Year1, Year2 As Integer

  If Data1 >Data2 Then
    DtSwap = Data1
    Data1 = Data2
    Data2 = DtSwap
  End If

  Year1 = Val(Format(Data1, "yyyy"))
  Month1 = Val(Format(Data1, "mm"))
  Day1 = Val(Format(Data1, "dd"))

  Year2 = Val(Format(Data2, "yyyy"))
  Month2 = Val(Format(Data2, "mm"))
  Day2 = Val(Format(Data2, "dd"))

  Anos = Year2 - Year1
  Meses = 0
  Dias = 0
  If Month2 <Month1 Then
    Meses = Meses + 12
    Anos = Anos -1
  End If
  Meses = Meses + (Month2 - Month1)
  If Day2 <Day1 Then
    Dias = Dias + DiasPorMes(Year1, Val(Month1))
    If Meses = 0 Then
      Anos = Anos -1
      Meses = 11
    Else
      Meses = Meses -1
    End If
  End If
  Dias = Dias + (Day2 - Day1)
End Sub

Function DiasPorMes(ByVal Ano, Mes As Integer)As Integer
  Dim Meses31 As String
  Dim Meses30 As String

  Meses31 = "'1','3','5','7','8','10','12'"
  Meses30 = "'4','6','9','11'"

  If InStr(Meses31, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 31
  ElseIf InStr(Meses30, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 30
  Else
    If Ano Mod 4 = 0 Then
      DiasPorMes = 29
    Else
      DiasPorMes = 28
    End If
  End If

End Function

Public Function ProcuraEventoAtivoInativo (pUltimoNivel As Boolean, TextoEvento As String, pTabelPreco As Long) As Long

  Dim sql As Object
  Set sql=NewQuery
  On Error GoTo Pula1
  sql.Clear
  sql.Add("SELECT HANDLE FROM SAM_TGE WHERE ESTRUTURANUMERICA=:P1")

  If pUltimoNivel Then
    sql.Add(" AND ULTIMONIVEL = 'S' ")
  End If
  If pTabelPreco > 0 Then
    sql.Add(" AND SAM_TGE.HANDLE IN (SELECT SAM_TGE_TABELATISS.EVENTO FROM SAM_TGE_TABELATISS WHERE SAM_TGE_TABELATISS.TABELATISS = " + CStr(pTabelPreco) + ")")
  End If

  sql.ParamByName("P1").AsInteger= CLng(TextoEvento)
  sql.Active=True
  If Not sql.EOF Then
    Dim sql2 As Object
    Set sql2 = NewQuery
    sql2.Clear
    sql2.Add("SELECT count(1) QTD FROM SAM_TGE WHERE ESTRUTURANUMERICA=:P1")
    sql2.ParamByName("P1").AsString=CLng(TextoEvento)
    If pUltimoNivel Then
      sql2.Add(" AND ULTIMONIVEL = 'S' ")
    End If
    If pTabelPreco > 0 Then
      sql2.Add(" AND SAM_TGE.HANDLE IN (SELECT SAM_TGE_TABELATISS.EVENTO FROM SAM_TGE_TABELATISS WHERE SAM_TGE_TABELATISS.TABELATISS = " + CStr(pTabelPreco) + ")")
    End If
    sql2.Active = True

	If sql2.FieldByName("QTD").AsInteger = 1 Then
      If sql.FieldByName("HANDLE").AsInteger>0 Then
        ProcuraEventoAtivoInativo=sql.FieldByName("HANDLE").AsInteger
        Set sql=Nothing
        Set sql2=Nothing
        Exit Function
      End If
    End If
    Set sql2=Nothing
  End If

  Pula1:
  On Error GoTo Pula2
  sql.Clear
  sql.Add("SELECT HANDLE FROM SAM_TGE WHERE ESTRUTURA=:P1")

  If pUltimoNivel Then
    sql.Add(" AND ULTIMONIVEL = 'S' ")
  End If
  If pTabelPreco > 0 Then
      sql.Add(" AND SAM_TGE.HANDLE IN (SELECT SAM_TGE_TABELATISS.EVENTO FROM SAM_TGE_TABELATISS WHERE SAM_TGE_TABELATISS.TABELATISS = " + CStr(pTabelPreco) + ")")
  End If

  sql.ParamByName("P1").AsString=TextoEvento
  sql.Active=True
  If Not sql.EOF Then
    Dim sql3 As Object
    Set sql3 = NewQuery
    sql3.Clear
    sql3.Add("SELECT count(1) QTD FROM SAM_TGE WHERE ESTRUTURA=:P1")
    sql3.ParamByName("P1").AsString=TextoEvento
    If pUltimoNivel Then
      sql3.Add(" AND ULTIMONIVEL = 'S' ")
    End If
    If pTabelPreco > 0 Then
      sql3.Add(" AND SAM_TGE.HANDLE IN (SELECT SAM_TGE_TABELATISS.EVENTO FROM SAM_TGE_TABELATISS WHERE SAM_TGE_TABELATISS.TABELATISS = " + CStr(pTabelPreco) + ")")
    End If
    sql3.Active = True

	If sql3.FieldByName("QTD").AsInteger = 1 Then
      If sql.FieldByName("HANDLE").AsInteger > 0 Then
        ProcuraEventoAtivoInativo = sql.FieldByName("HANDLE").AsInteger
        Set sql = Nothing
        Set sql3 = Nothing
        Exit Function
      End If
    End If
    Set sql3 = Nothing
  End If
  Pula2:
  Set sql=Nothing
  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vOrdem As Integer

  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_TGE.ESTRUTURA|SAM_TGE.Z_DESCRICAO|SAM_TGE.DESCRICAOABREVIADA|SAM_TGE.NIVELAUTORIZACAO"

  If pUltimoNivel Then
    vCriterio = "SAM_TGE.ULTIMONIVEL = 'S' "
  Else
    vCriterio = "SAM_TGE.HANDLE > 0"
  End If

  vCriterio = vCriterio + " AND (SAM_TGE.INATIVO = 'S' OR SAM_TGE.INATIVO = 'N') "

  If pTabelPreco > 0 Then
      vCriterio = vCriterio + " AND SAM_TGE.HANDLE IN (SELECT SAM_TGE_TABELATISS.EVENTO FROM SAM_TGE_TABELATISS WHERE SAM_TGE_TABELATISS.TABELATISS = " + CStr(pTabelPreco) + ")"
  End If

  If IsInt(TiraAcento(TextoEvento,True)) Then
    vOrdem = 2
  Else
    vOrdem = 3
  End If


  vCampos = "Estrutura TGE|Descrição TGE|Descrição abreviada TGE|Nível"

  ProcuraEventoAtivoInativo = Interface.Exec(CurrentSystem, "SAM_TGE",vColunas,vOrdem ,vCampos,vCriterio, _
  "Tabela Geral de Eventos",False,TextoEvento,"CA011.ConsultaTge")

  Set Interface = Nothing

End Function

Public Function VerificaAgrupadorPagamentoFechado As Boolean
	Dim qPeg As Object
	Set qPeg = NewQuery
	qPeg.Add("SELECT G.PEG FROM SAM_GUIA G WHERE G.HANDLE = :HANDLE")
	qPeg.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIA").AsInteger
	qPeg.Active = True
	Dim callEntity As CSEntityCall
  	Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamPeg, Benner.Saude.Entidades", "VerificaPegVinculadoPagamentoFechado")
  	callEntity.AddParameter(pdtAutomatic, qPeg.FieldByName("PEG").AsInteger)
  	VerificaAgrupadorPagamentoFechado = CBool(callEntity.Execute)
	Set callEntity =  Nothing
	Set qPeg = Nothing
End Function

Public Sub MeuBeforeEdit(ShowPopup As Boolean)

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.SamGuiaEventos")

	vDllBSPro006.MeuBeforeEdit(CurrentSystem, _
	                           CurrentQuery.TQuery, _
							   ShowPopup)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub VERIFICAJAPAGO(CanContinue As Boolean)
  Dim SITUACAO As String
  Dim Q1 As Object
  Set Q1 = NewQuery
  Q1.Clear
  Q1.Add("SELECT G.SITUACAO FROM SAM_GUIA G WHERE G.HANDLE=:GUIA")
  Q1.ParamByName("GUIA").Value = RecordHandleOfTableInterfacePEG("SAM_GUIA")
  Q1.Active = True
  SITUACAO = Q1.FieldByName("SITUACAO").AsString
  Q1.Active = False
  Set Q1 = Nothing
  If SITUACAO = "4" Then
    CanContinue = False
  Else
    CanContinue = True
  End If
End Sub

Public Sub SugerirIdadeBeneficiario
  'sugerir a idade Do BENEFICIARIO
  Dim vAux As Long
  Dim database As Date

  Dim qParamGeral As Object
  Set qParamGeral = NewQuery

  qParamGeral.Clear
  qParamGeral.Add("SELECT SUGERIRIDADEBENEF FROM SAM_PARAMETROSPROCCONTAS")
  qParamGeral.Active = True


  If qParamGeral.FieldByName("SUGERIRIDADEBENEF").AsString = "S" Then
    If CurrentQuery.State <>1 Then
      If CurrentQuery.FieldByName("DATAATENDIMENTO").IsNull Then
        database = ServerDate
      Else
        database = CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime
      End If
      vAux = CalculaIdadeBeneficiario(CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, database)
      If vAux >= 0 And vAux <>CurrentQuery.FieldByName("IDADEBENEFICIARIO").AsInteger Then
        CurrentQuery.FieldByName("IDADEBENEFICIARIO").Value = vAux
      End If
    End If
  End If

  Set qParamGeral = Nothing

End Sub
