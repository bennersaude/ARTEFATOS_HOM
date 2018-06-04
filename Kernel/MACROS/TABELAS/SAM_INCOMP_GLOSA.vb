'HASH: 6710491979A18FF11FAAE13CCD506625
'Macro: SAM_INCOMP_GLOSA

'#Uses "*bsShowMessage
'#Uses "*RecordHandleOfTableInterfacePEG"
'#Uses "*RefreshNodesWithTableInterfacePEG"

Option Explicit

Dim vDllBsInterface0047 As Object
Dim vDllBSPro006        As Object

Dim qConsulta           As Object

Public Sub BOTAOAMBOS_OnClick()
  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  If aux = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("SamIncomp.Check")

  Interface.ProntuarioBeneficiario(CurrentSystem, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "A")

  Set Interface = Nothing
End Sub

Public Sub BOTAOHISTBUCAL_OnClick()
  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  If aux = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  Dim HistoricoDll As Object
  Dim vAutorizacao As Integer
  Dim vSituacaoAnteriorGuia As String
  Dim vGuia As Long
  Dim qGuia As Object

  Set qGuia = NewQuery

  qGuia.Active = False

  qGuia.Clear
  qGuia.Add("SELECT SG.HANDLE GUIA,")
  qGuia.Add("       SG.SITUACAO SITUACAO")
  qGuia.Add("  FROM SAM_GUIA_EVENTOS SGE, ")
  qGuia.Add("       SAM_GUIA SG ")
  qGuia.Add(" WHERE SGE.HANDLE = :GUIAEVENTO ")
  qGuia.Add("   AND SGE.GUIA = SG.HANDLE ")
  qGuia.ParamByName("GUIAEVENTO").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger

  qGuia.Active = True

  'Se estiver em Pronto,permite acessar o Historico bucal
  If qGuia.FieldByName("SITUACAO").AsString = "3" Then
    vSituacaoAnteriorGuia = qGuia.FieldByName("SITUACAO").AsString
    vGuia = qGuia.FieldByName("GUIA").AsInteger
    vAutorizacao = 0

    Set HistoricoDll = CreateBennerObject("BSCLI006.ROTINAS")
    HistoricoDll.HistoricoBucal(CurrentSystem, vGuia, vAutorizacao)

    Set HistoricoDll = Nothing

    'Verifica se a guia foi reprocessada
    qGuia.Active = False

    qGuia.Clear
    qGuia.Add("SELECT SG.SITUACAO SITUACAO")
    qGuia.Add("  FROM SAM_GUIA_EVENTOS SGE, ")
    qGuia.Add("       SAM_GUIA SG ")
    qGuia.Add(" WHERE SGE.HANDLE = :GUIAEVENTO ")
    qGuia.Add("   AND SGE.GUIA = SG.HANDLE ")
    qGuia.ParamByName("GUIAEVENTO").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger

    qGuia.Active = True

    If(vSituacaoAnteriorGuia <>qGuia.FieldByName("SITUACAO").AsString)Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")
  End If
  'Fim verificação
Else
  bsShowMessage("A guia original deve estar em Pronto!", "E")

End If

Set qGuia = Nothing
End Sub

Public Sub BOTAOLIBERAR_OnClick()
Dim Interface As Object
Dim aux As Boolean
Dim vsresultado As String
Dim vbContinua As Boolean

  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  If aux = False Then
    If VisibleMode Then
      RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    End If
    Exit Sub
  End If

  If(CurrentQuery.State <>1)Then
    bsShowMessage("A incompatibilidade não pode estar em edição", "E")
    Exit Sub
  End If

  If(CurrentQuery.FieldByName("SITUACAO").AsString <>"P")Then
    bsShowMessage("A incompatibilidade não está pendente", "E")
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SamAcertos.Incompatibilidade")
  Interface.Consistencia(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsresultado)

  vbContinua = False

  If vsresultado <> "" Then
    If bsShowMessage(vsresultado + Chr(13) + "Deseja Continuar ?", "Q") = vbYes Then
      vbContinua = True
    End If
  Else
    vbContinua = True
  End If

  If (Not vbContinua) Then
    Exit Sub
  End If

  '//SMS 37769 em 03/06/2005 - Wagner
  'Incluídos os campos: USUARIOLIBERACAO e DATAHORALIBERACAO
  StartTransaction
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("UPDATE SAM_INCOMP_GLOSA SET SITUACAO='L',USUARIOPROCESSAMENTO=:U, DATAPROCESSAMENTO=:D, USUARIOLIBERACAO=:U, DATAHORALIBERACAO=:D WHERE HANDLE=:H")
  SQL.ParamByName("U").AsInteger = CurrentUser
  SQL.ParamByName("D").AsDateTime = ServerNow
  SQL.ParamByName("H").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ExecSQL
  Commit
  CurrentQuery.Active = False
  CurrentQuery.Active = True

  If VisibleMode Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")
  End If

  'Grava registro na auditoria Do sistema indicando qual responsável (código) que fez a alteração.
  WriteAudit("A", HandleOfTable("SAM_INCOMP_GLOSA"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Usuário: " + Str(CurrentUser) + " - Data/hora: " + Str(ServerNow) + " - Observação: " + CurrentQuery.FieldByName("OBSERVACAOLIBERACAO").AsString)

End Sub

Public Sub BOTAOMEDICO_OnClick()
  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  If aux = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("SamIncomp.Check")

  Interface.ProntuarioBeneficiario(CurrentSystem, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "M")

  Set Interface = Nothing

End Sub

Public Sub BOTAOODONTO_OnClick()
  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  If aux = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface = CreateBennerObject("SamIncomp.Check")

  Interface.ProntuarioBeneficiario(CurrentSystem, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "O")

  Set Interface = Nothing

End Sub

Public Sub MOTIVOGLOSAANTERIOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim Interface As Object
  ShowPopup = False

  vColunas = "CODIGOGLOSA|SAM_MOTIVOGLOSA.DESCRICAO|SAM_TIPOMOTIVOGLOSA.DESCRICAO"
  vCampos = "Código|Descrição|Tipo Glosa"
  vCriterio = "SAM_MOTIVOGLOSA.ATIVA='S'"

  Set Interface = CreateBennerObject("Procura.Procurar")
  vHandle = Interface.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE]", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, "")
  Set Interface = Nothing

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").Value = vHandle
  End If
End Sub

Public Sub MOTIVOGLOSAPOSTERIOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim Interface As Object
  ShowPopup = False

  vColunas = "CODIGOGLOSA|SAM_MOTIVOGLOSA.DESCRICAO|SAM_TIPOMOTIVOGLOSA.DESCRICAO"
  vCampos = "Código|Descrição|Tipo Glosa"
  vCriterio = "SAM_MOTIVOGLOSA.ATIVA='S'"

  Set Interface = CreateBennerObject("Procura.Procurar")
  vHandle = Interface.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE]", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, "")
  Set Interface = Nothing

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").Value = vHandle
  End If
End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    MOTIVOGLOSAANTERIOR.WebLocalWhere = "A.ATIVA='S'"
    MOTIVOGLOSAPOSTERIOR.WebLocalWhere = "A.ATIVA='S'"
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    Dim qAux As Object
    Set qAux = NewQuery

    qAux.Active = False

    qAux.Clear
    qAux.Add("SELECT PARAM.UTILIZAHISTORICOODONTOLOGICO UTILIZHISTODONTO, ")
    qAux.Add("       SP.EXIGEHISTORICOBUCAL EXIGEHISTBUCALPRES, ")
    qAux.Add("       MDGUIA.EXIGEHISTORICOBUCAL EXIGEHISTBUCALMODGUIA")
    qAux.Add("  FROM SAM_PARAMETROSPROCCONTAS PARAM, ")
    qAux.Add("       SAM_PRESTADOR SP, ")
    qAux.Add("       SAM_TIPOGUIA_MDGUIA	MDGUIA,")
    qAux.Add("       SAM_TIPOGUIA ST, ")
    qAux.Add("       SAM_GUIA_EVENTOS SGE, ")
    qAux.Add("       SAM_GUIA SG, ")
    qAux.Add("       SAM_PEG PEG ")
    qAux.Add(" WHERE SGE.HANDLE = :GUIAEVENTO ")
    qAux.Add("   AND SGE.GUIA = SG.HANDLE ")
    qAux.Add("   AND SG.MODELOGUIA = MDGUIA.HANDLE ")
    qAux.Add("   AND MDGUIA.TIPOGUIA = ST.HANDLE ")
    qAux.Add("   AND ST.TABTIPOGUIA = 3 ")

    ' Luciano T. Alberti - SMS 65194 - 14/07/2006 - Início
    ' >>>> Melhoria de desempenho, pois o OR estava causando queda de performace

    'qAux.Add("   AND ((SG.EXECUTOR = SP.HANDLE AND SG.EXECUTOR IS NOT NULL) ")
    'qAux.Add("        OR (SG.PEG = PEG.HANDLE AND PEG.LOCALEXECUCAO = SP.HANDLE))")

    qAux.Add("   AND (SG.EXECUTOR = SP.HANDLE AND SG.EXECUTOR IS NOT NULL) ")
    qAux.Add("UNION ALL")
    qAux.Add("SELECT PARAM.UTILIZAHISTORICOODONTOLOGICO UTILIZHISTODONTO, ")
    qAux.Add("       SP.EXIGEHISTORICOBUCAL EXIGEHISTBUCALPRES, ")
    qAux.Add("       MDGUIA.EXIGEHISTORICOBUCAL EXIGEHISTBUCALMODGUIA")
    qAux.Add("  FROM SAM_PARAMETROSPROCCONTAS PARAM, ")
    qAux.Add("       SAM_PRESTADOR SP, ")
    qAux.Add("       SAM_TIPOGUIA_MDGUIA	MDGUIA,")
    qAux.Add("       SAM_TIPOGUIA ST, ")
    qAux.Add("       SAM_GUIA_EVENTOS SGE, ")
    qAux.Add("       SAM_GUIA SG, ")
    qAux.Add("       SAM_PEG PEG ")
    qAux.Add(" WHERE SGE.HANDLE = :GUIAEVENTO ")
    qAux.Add("   AND SGE.GUIA = SG.HANDLE ")
    qAux.Add("   AND SG.MODELOGUIA = MDGUIA.HANDLE ")
    qAux.Add("   AND MDGUIA.TIPOGUIA = ST.HANDLE ")
    qAux.Add("   AND ST.TABTIPOGUIA = 3 ")
    qAux.Add("   AND (SG.PEG = PEG.HANDLE AND PEG.LOCALEXECUCAO = SP.HANDLE)")
    ' Luciano T. Alberti - SMS 65194 - 14/07/2006 - fIM

    qAux.ParamByName("GUIAEVENTO").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger

    qAux.Active = True

    If(Not(qAux.FieldByName("UTILIZHISTODONTO").AsString = "S")And Not(qAux.FieldByName("EXIGEHISTBUCALPRES").AsString = "S") _
       And Not(qAux.FieldByName("EXIGEHISTBUCALMODGUIA").AsString = "S"))Then
    BOTAOHISTBUCAL.Visible = False
  Else
    BOTAOHISTBUCAL.Visible = True
  End If
End If
'SMS 54845 - Marcelo Barbosa - 23/02/2006
'Conforme SMS é para deixar apenas o campo observação deverá ficar livre para edição
'//SMS 37769 em 03/06/2005 - Wagner
'If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    CONSIDERAEXECUTOR.ReadOnly = True
    CONSIDERALOCALEXECUCAO.ReadOnly = True
    CONSIDERACBOCBOS.ReadOnly = True
    CONSIDERARECEBEDOR.ReadOnly = True
    MOTIVOGLOSAANTERIOR.ReadOnly = True
    PERCENTGLOSAANTERIOR.ReadOnly = True
    MOTIVOGLOSAPOSTERIOR.ReadOnly = True
    PERCENTGLOSAPOSTERIOR.ReadOnly = True
'Else
'  CONSIDERAEXECUTOR.ReadOnly = False
'  MOTIVOGLOSAANTERIOR.ReadOnly = False
'  PERCENTGLOSAANTERIOR.ReadOnly = False
'  MOTIVOGLOSAPOSTERIOR.ReadOnly = False
'  PERCENTGLOSAPOSTERIOR.ReadOnly = False
'End If
'Fim - SMS 54845
USUARIOLIBERACAO.ReadOnly = True
DATAHORALIBERACAO.ReadOnly = True
OBSERVACAOLIBERACAO.ReadOnly = False

  'Luciano T. Alberti - SMS 65194 - 14/07/2006 - Deve permitir alterar apenas se for do tipo AUDITAR
  If CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull  And _
     CurrentQuery.FieldByName("TABTIPOACAO").AsInteger = 2 Then
    CONSIDERAEXECUTOR.ReadOnly = False
    CONSIDERALOCALEXECUCAO.ReadOnly = False
    CONSIDERACBOCBOS.ReadOnly = False
    CONSIDERARECEBEDOR.ReadOnly = False
    MOTIVOGLOSAANTERIOR.ReadOnly = False
    PERCENTGLOSAANTERIOR.ReadOnly = False
    MOTIVOGLOSAPOSTERIOR.ReadOnly = False
    PERCENTGLOSAPOSTERIOR.ReadOnly = False
  End If
  'Luciano T. Alberti - SMS 65194 - 14/07/2006 - Fim


End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  CanContinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "E")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  CanContinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  CanContinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "I")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'SMS 74956 - Marcelo Barbosa - 10/01/2007
  Dim qParam As Object

  Set qParam = NewQuery

  qParam.Clear
  qParam.Add("SELECT USARPRONTUARIO FROM SAM_PARAMETROSPROCCONTAS")
  qParam.Active = True

  If qParam.FieldByName("USARPRONTUARIO").AsString <> "S" Then
    If CurrentQuery.FieldByName("COMPETENCIAANTERIOR").IsNull Then
      CanContinue = False
      bsShowMessage("Campo Competência Anterior é obrigatório!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("EVENTOGUIAANTERIOR").IsNull Then
      CanContinue = False
      bsShowMessage("Campo Evento Anterior é obrigatório!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("GUIAANTERIOR").IsNull Then
      CanContinue = False
      bsShowMessage("Campo Guia Anterior é obrigatório!", "E")
      Exit Sub
    End If
    If Len(ORDEMEVENTOGUIAANTERIOR.Text) = 0 Then
      CanContinue = False
      bsShowMessage("Campo Ordem do Evento da Guia Anterior é obrigatório!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("PEGANTERIOR").IsNull Then
      CanContinue = False
      bsShowMessage("Campo PEG Anterior é obrigatório!", "E")
      Exit Sub
    End If
    If CurrentQuery.FieldByName("VALORANTERIOR").IsNull Then
      CanContinue = False
      bsShowMessage("Campo Valor Anterior é obrigatório!", "E")
      Exit Sub
    End If
  End If
  'Fim - SMS 74956

  If(CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat <>0)And(CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull)Then
    CanContinue = False
    bsShowMessage("Motivo de glosa do evento anterior é obrigatório quando percentual da glosa é diferente de zero", "E")
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
         Case "BOTAOLIBERAR"
			 BOTAOLIBERAR_OnClick
 		 Case "BOTAOGLOSAR"
			 BOTAOGLOSAR_OnClick
 End Select

End Sub

Public Sub TABTIPOACAO_OnChanging(AllowChange As Boolean)
  bsShowMessage("Alteração não permitida", "E")
  AllowChange = False
End Sub

Public Function JAPAGO(guiaevento As Long)As Boolean
  Dim SITUACAO As String
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Clear
  'q1.Add("SELECT G.SITUACAO FROM SAM_GUIA G JOIN SAM_GUIA_EVENTOS E ON (E.GUIA=G.HANDLE) WHERE E.HANDLE=:GUIAEVENTO")
  q1.Add("SELECT G.SITUACAO")
  q1.Add("  FROM SAM_GUIA G, ")
  q1.Add("       SAM_GUIA_EVENTOS")
  q1.Add(" WHERE E.HANDLE = :GUIAEVENTO")
  q1.Add("   AND E.GUIA   = G.HANDLE")
  q1.ParamByName("GUIAEVENTO").Value = guiaevento
  q1.Active = True
  SITUACAO = q1.FieldByName("SITUACAO").AsString
  q1.Active = False
  Set q1 = Nothing
  If SITUACAO = "4" Then
    JAPAGO = True
  Else
    JAPAGO = False
  End If
End Function

Public Sub BOTAOGLOSAR_OnClick()

Dim Interface As Object
Dim giEventoAnteriorPago As Long
Dim giEventoPosteriorPago As Long
Dim viEventoAnteriorGlosado As Long
Dim viEventoPosteriorGlosado As Long
Dim viRetorno As Long
Dim VsResultado As String
Dim vsTipoMensagem As String
Dim vsmensagem As String
Dim vbContinua As Boolean

'If VisibleMode Then
'  Dim vsMensagemErro As String
'
'  Set Interface = CreateBennerObject("BSINTERFACE0002.GERARFORMULARIOVIRTUAL")
'  Interface.Exec(CurrentSystem, _
'                 1, _
'                 "TV_FORM0052", _
'                 "Glosa do Evento Anterior", _
'                  0, _
'                  500, _
'                  633, _
'                  False, _
'                  vsMensagemErro, _
'                  "")
'End If

  If (WebMode) Then

	Dim HandleIncompGlosa As Long
	HandleIncompGlosa = RecordHandleOfTableInterfacePEG("SAM_INCOMP_GLOSA")
	Dim qImcompGlosa As Object
	Set qImcompGlosa = NewQuery

	qImcompGlosa.Clear
	qImcompGlosa.Add("SELECT *")
	qImcompGlosa.Add("  FROM SAM_INCOMP_GLOSA ")
	qImcompGlosa.Add(" WHERE HANDLE = :HANDLE")
	qImcompGlosa.ParamByName("HANDLE").Value = HandleIncompGlosa
	qImcompGlosa.Active = True

	If(qImcompGlosa.FieldByName("SITUACAO").AsString <>"P")Then
	bsShowMessage("A incompatibilidade não está pendente", "I")
	Set qImcompGlosa = Nothing
	Exit Sub
	End If

	If(qImcompGlosa.FieldByName("EVENTOGUIAANTERIOR").IsNull And _
	qImcompGlosa.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And _
	qImcompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").IsNull And _
	qImcompGlosa.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull)Then
	bsShowMessage("Não há eventos a serem glosados", "E")
	Set qImcompGlosa = Nothing
	Exit Sub
	Else
	If(qImcompGlosa.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 0)And _
	 (qImcompGlosa.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 0)Then
	bsShowMessage("Não existe definição de Percentual de Glosa", "E")
	Set qImcompGlosa = Nothing
	Exit Sub
	End If
	End If

	Set Interface = CreateBennerObject("SamAcertos.Incompatibilidade")
	Interface.Consistencia(CurrentSystem, HandleIncompGlosa, VsResultado)
	viEventoAnteriorGlosado = 0
	viEventoPosteriorGlosado  = 0

	If VsResultado <> "" Then
	bsShowMessage(VsResultado, "I")
	End If

	vsmensagem = Interface.VerificaPendenciasEventoAnterior(CurrentSystem, HandleIncompGlosa, vsTipoMensagem, viEventoAnteriorGlosado)
	If (vsmensagem <> "") Then
	bsShowMessage(vsmensagem, "I")
	End If
	vsmensagem = Interface.VerificaPendenciasEventoPosterior(CurrentSystem, HandleIncompGlosa, vsTipoMensagem, viEventoPosteriorGlosado)
	If (vsmensagem <> "") Then
	bsShowMessage(vsmensagem, "I")
	End If

	giEventoAnteriorPago = 0
	giEventoPosteriorPago = 0

	vsmensagem = ""
	viRetorno = Interface.Glosar(CurrentSystem, _
						     HandleIncompGlosa, _
	                         giEventoAnteriorPago, _
						     giEventoPosteriorPago, _
	                         viEventoAnteriorGlosado, _
						     viEventoPosteriorGlosado, _
	                         vsmensagem)

	If (viRetorno = 0) Then
	bsShowMessage(vsmensagem + "Impossível Continuar, Verifique os erros!", "E")
	vbContinua = False
	ElseIf (viRetorno = 1) And (Trim(vsmensagem) <> "") Then
	bsShowMessage(vsmensagem, "I")
	If (viEventoAnteriorGlosado = 1) And (viEventoPosteriorGlosado = 1) Then
	Set qImcompGlosa = Nothing
	Exit Sub
	End If
	End If


	Dim qPegGuia As Object
	Set qPegGuia = NewQuery

	If (viRetorno = 1) And (giEventoAnteriorPago = 1) Then
	qPegGuia.Active = False
	qPegGuia.Clear
	qPegGuia.Add("SELECT SOFREUREAPRESENTACAO        ")
	qPegGuia.Add("  FROM SAM_GUIA_EVENTOS            ")
	qPegGuia.Add(" WHERE HANDLE = :EVENTOGUIAANTERIOR")
	qPegGuia.ParamByName("EVENTOGUIAANTERIOR").AsInteger = qImcompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
	qPegGuia.Active = True

	If qPegGuia.FieldByName("SOFREUREAPRESENTACAO").AsString <> "S" Then
	qPegGuia.Active = False
	qPegGuia.Clear
	qPegGuia.Add("SELECT P.PEG, G.GUIA                             ")
	qPegGuia.Add("  FROM SAM_PEG P                                 ")
	qPegGuia.Add("  JOIN SAM_GUIA G ON G.PEG = P.HANDLE            ")
	qPegGuia.Add("  JOIN SAM_GUIA_EVENTOS GE ON GE.GUIA = G.HANDLE ")
	qPegGuia.Add(" WHERE GE.HANDLE = :EVENTOGUIAANTERIOR           ")
	qPegGuia.ParamByName("EVENTOGUIAANTERIOR").AsInteger = qImcompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
	qPegGuia.Active = True

	BsShowMessage("Evento anterior pago. Necessário realizar a reapresentação de eventos do PEG " + CStr(qPegGuia.FieldByName("PEG").AsInteger) + " e Guia " + CStr(qPegGuia.FieldByName("GUIA").AsInteger), "I")
	End If


	ElseIf (viRetorno = 1) And (giEventoPosteriorPago = 1) Then
	qPegGuia.Active = False
	qPegGuia.Clear
	qPegGuia.Add("SELECT SOFREUREAPRESENTACAO        ")
	qPegGuia.Add("  FROM SAM_GUIA_EVENTOS            ")
	qPegGuia.Add(" WHERE HANDLE = :EVENTOGUIAPOSTERIOR")
	qPegGuia.ParamByName("EVENTOGUIAPOSTERIOR").AsInteger = qImcompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger
	qPegGuia.Active = True

	If qPegGuia.FieldByName("SOFREUREAPRESENTACAO").AsString <> "S" Then
	  qPegGuia.Active = False
	  qPegGuia.Clear
	  qPegGuia.Add("SELECT P.PEG, G.GUIA                             ")
	  qPegGuia.Add("  FROM SAM_PEG P                                 ")
	  qPegGuia.Add("  JOIN SAM_GUIA G ON G.PEG = P.HANDLE            ")
	  qPegGuia.Add("  JOIN SAM_GUIA_EVENTOS GE ON GE.GUIA = G.HANDLE ")
	  qPegGuia.Add(" WHERE GE.HANDLE = :EVENTOGUIAPOSTERIOR           ")
	  qPegGuia.ParamByName("EVENTOGUIAPOSTERIOR").AsInteger = qImcompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger
	  qPegGuia.Active = True

	  BsShowMessage("Evento posterior pago. Necessário realizar a reapresentação de eventos do PEG " + CStr(qPegGuia.FieldByName("PEG").AsInteger) + " e Guia " + CStr(qPegGuia.FieldByName("GUIA").AsInteger), "I")
	End If
	End If


	Set qPegGuia = Nothing
	Set qImcompGlosa = Nothing

  Else

    Dim aux As Boolean
	aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTableInterfacePEG("FILIAIS"), "A")
	If aux = False Then
	  RefreshNodesWithTableInterfacePEG("SAM_INCOMP_GLOSA")
	  Exit Sub
	End If

    If(CurrentQuery.State <>1)Then
	  bsShowMessage("A incompatibilidade não pode estar em edição", "E")
	  Exit Sub
	End If

    If(CurrentQuery.FieldByName("SITUACAO").AsString <>"P")Then
	  bsShowMessage("A incompatibilidade não está pendente", "I")
	  Exit Sub
	End If

    Dim consistencia As Long

    If(CurrentQuery.FieldByName("EVENTOGUIAANTERIOR").IsNull And _
	   CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And _
	   CurrentQuery.FieldByName("EVENTOGUIAPOSTERIOR").IsNull And _
	   CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull)Then
      bsShowMessage("Não há eventos a serem glosados", "I")
	Else
	  If(CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 0)And _
	     (CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 0)Then
	    bsShowMessage("Não existe definição de Percentual de Glosa", "E")
	    Exit Sub
	  End If
	End If

	Set Interface = CreateBennerObject("SamAcertos.Incompatibilidade")
	Interface.Consistencia(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, VsResultado)

	viEventoAnteriorGlosado = 0
	viEventoPosteriorGlosado  = 0

	vbContinua = False

	If VsResultado <> "" Then
	  If bsShowMessage(VsResultado + Chr(13) + "Deseja Continuar ?", "Q") = vbYes Then
	    vbContinua = True
	  End If
	Else
	  vbContinua = True
	End If

	If vbContinua Then
	  vsmensagem = Interface.VerificaPendenciasEventoAnterior(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsTipoMensagem, viEventoAnteriorGlosado)
	  If (vsmensagem <> "") Then
	    If (vsTipoMensagem = "I") Then
	      bsShowMessage(vsmensagem,"I")
	    ElseIf (vsTipoMensagem = "Q") Then
	      vbContinua = False
	      If bsShowMessage(vsmensagem + ". Continuar assim mesmo ?", vsTipoMensagem) = vbYes Then
	        vbContinua = True
	      End If
	    End If
	  End If
	End If

	If vbContinua Then
	  vsmensagem = Interface.VerificaPendenciasEventoPosterior(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vsTipoMensagem, viEventoPosteriorGlosado)
	  If (vsmensagem <> "") Then
	    If (vsTipoMensagem = "I") Then
	      bsShowMessage(vsmensagem,"I")
	    ElseIf (vsTipoMensagem = "Q") Then
	      vbContinua = False
	      If bsShowMessage(vsmensagem + ". Continuar assim mesmo ?", vsTipoMensagem) = vbYes Then
	        vbContinua = True
	      End If
	    End If
	  End If
	End If

    giEventoAnteriorPago = 0
    giEventoPosteriorPago = 0

    If vbContinua Then
      vsmensagem = ""
      viRetorno = Interface.Glosar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, _
                             giEventoAnteriorPago, giEventoPosteriorPago, _
                             viEventoAnteriorGlosado, viEventoPosteriorGlosado, _
                             vsmensagem)

	    If (viRetorno = 0) Then
	  	  bsShowMessage(vsmensagem + "Impossível Continuar, Verifique os erros!", "E")
		  vbContinua = False
	    ElseIf (viRetorno = 1) And (Trim(vsmensagem) <> "") Then
		  bsShowMessage(vsmensagem, "I")
	    End If
    End If
    Set Interface = Nothing
  End If

End Sub

Public Sub K9BOTAOALERTAANTERIOR_OnClick()

	Set vDllBsInterface0047 = CreateBennerObject("BSINTERFACE0047.Rotinas")

	vDllBsInterface0047.VerAlertas(CurrentSystem, CurrentQuery.FieldByName("EVENTOGUIAANTERIOR").AsInteger)

	Set vDllBsInterface0047 = Nothing

End Sub

Public Sub K9BOTAOALERTAPOSTERIOR_OnClick()

	Set vDllBsInterface0047 = CreateBennerObject("BSINTERFACE0047.Rotinas")

	vDllBsInterface0047.VerAlertas(CurrentSystem, CurrentQuery.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger)

	Set vDllBsInterface0047 = Nothing

End Sub

Public Sub K9BOTAOEVENTOANTEIROR_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

	vDllBSPro006.Detalhe(CurrentSystem, "SAM_GUIA_EVENTOS", CurrentQuery.FieldByName("EVENTOGUIAANTERIOR").AsInteger)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub K9BOTAOEVENTOPOSTERIOR_OnClick()

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

	vDllBSPro006.Detalhe(CurrentSystem, "SAM_GUIA_EVENTOS", CurrentQuery.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger)

	Set vDllBSPro006 = Nothing

End Sub

Public Sub K9BOTAOEXECANTEIROR_OnClick()

	Dim gDllCa005 As Object

	Set gDllCa005 = CreateBennerObject ("CA005.ConsultaPrestador")

	gDllCa005.info(CurrentSystem, CurrentQuery.FieldByName("EXECUTORANTERIOR").AsInteger)

	Set gDllCa005 = Nothing

End Sub

Public Sub K9BOTAOEXECPOSTERIOR_OnClick()

	Dim gDllCa005 As Object

	Set gDllCa005 = CreateBennerObject ("CA005.ConsultaPrestador")

	gDllCa005.info(CurrentSystem, CurrentQuery.FieldByName("EXECUTORPOSTERIOR").AsInteger)

	Set gDllCa005 = Nothing

End Sub

Public Sub K9BOTAOGUIAANTERIOR_OnClick()

	Set qConsulta = NewQuery

	qConsulta.Active = False
	qConsulta.Clear
	qConsulta.Add("SELECT GUIA             ")
	qConsulta.Add("  FROM SAM_GUIA_EVENTOS ")
	qConsulta.Add(" WHERE HANDLE = :PHANDLE")
	qConsulta.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOGUIAANTERIOR").AsInteger
	qConsulta.Active = True

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

	vDllBSPro006.Detalhe(CurrentSystem, "SAM_GUIA", qConsulta.FieldByName("GUIA").AsInteger)

	Set vDllBSPro006 = Nothing

	Set qConsulta = Nothing

End Sub

Public Sub K9BOTAOGUIAPOSTERIOR_OnClick()

	Set qConsulta = NewQuery

	qConsulta.Active = False
	qConsulta.Clear
	qConsulta.Add("SELECT GUIA             ")
	qConsulta.Add("  FROM SAM_GUIA_EVENTOS ")
	qConsulta.Add(" WHERE HANDLE = :PHANDLE")
	qConsulta.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger
	qConsulta.Active = True

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

	vDllBSPro006.Detalhe(CurrentSystem, "SAM_GUIA", qConsulta.FieldByName("GUIA").AsInteger)

	Set vDllBSPro006 = Nothing

	Set qConsulta = Nothing

End Sub

Public Sub K9BOTAOPRECOANTEIROR_OnClick()

	Set qConsulta = NewQuery

	qConsulta.Active = False
	qConsulta.Clear
	qConsulta.Add("SELECT HANDLE                 ")
	qConsulta.Add("  FROM SAM_GUIA_EVENTOS_PRECO ")
	qConsulta.Add(" WHERE GUIAEVENTO = :PHANDLE  ")
	qConsulta.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOGUIAANTERIOR").AsInteger
	qConsulta.Active = True

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

	vDllBSPro006.Detalhe(CurrentSystem, "SAM_GUIA_EVENTOS_PRECO", qConsulta.FieldByName("HANDLE").AsInteger)

	Set vDllBSPro006 = Nothing

	Set qConsulta = Nothing

End Sub

Public Sub K9BOTAOPRECOPOSTERIOR_OnClick()

	Set qConsulta = NewQuery

	qConsulta.Active = False
	qConsulta.Clear
	qConsulta.Add("SELECT HANDLE                 ")
	qConsulta.Add("  FROM SAM_GUIA_EVENTOS_PRECO ")
	qConsulta.Add(" WHERE GUIAEVENTO = :PHANDLE  ")
	qConsulta.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger
	qConsulta.Active = True

	Set vDllBSPro006 = CreateBennerObject("BSPRO006.Rotinas")

	vDllBSPro006.Detalhe(CurrentSystem, "SAM_GUIA_EVENTOS_PRECO", qConsulta.FieldByName("HANDLE").AsInteger)

	Set vDllBSPro006 = Nothing

	Set qConsulta = Nothing

End Sub

Public Sub DetalheRecebedor(psCampo As String)

	Dim qGuiaEvento As Object
	Set qGuiaEvento = NewQuery

	qGuiaEvento.Active = False
	qGuiaEvento.Clear
	qGuiaEvento.Add("SELECT SGE.RECEBEDOR                 ")
	qGuiaEvento.Add("  FROM SAM_GUIA_EVENTOS SGE (NOLOCK) ")
	qGuiaEvento.Add(" WHERE SGE.HANDLE = :PHANDLE         ")
	qGuiaEvento.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName(psCampo).AsInteger
	qGuiaEvento.Active = True

	Dim gDllCa005 As Object

	Set gDllCa005 = CreateBennerObject ("CA005.ConsultaPrestador")

	gDllCa005.info(CurrentSystem, qGuiaEvento.FieldByName("RECEBEDOR").AsInteger)

	Set gDllCa005 = Nothing

	Set qGuiaEvento = Nothing

End Sub

Public Sub K9BOTAORECEBEDORANTERIOR_OnClick()

	DetalheRecebedor("EVENTOGUIAANTERIOR")

End Sub

Public Sub K9BOTAORECEBEDORPOSTERIOR_OnClick()

	DetalheRecebedor("EVENTOGUIAPOSTERIOR")

End Sub
