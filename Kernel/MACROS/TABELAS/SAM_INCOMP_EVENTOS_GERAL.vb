'HASH: D46940C3F46B585732DCF71F859EEF55
'Henrique 15/02/2002
'Função ProcuraGrausValido -para pegar somente os graus válidos para o evento
'Passar o evento e o Tipo: "A" -Anterior;"P" -Posterior
'Caso já tiver uma inconpatibilidade cadastrada ñ deixa cadastrar
'Alterado por: Soares - SMS: 60815 - 23/05/2006
'#Uses "*bsShowMessage"
'#Uses "*ProcuraEvento"

Option Explicit

Public Sub EVENTOPOSTERIOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOPOSTERIOR.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOPOSTERIOR").Value = vHandle
  End If

End Sub

Public Sub GRAUANTERIOR_OnPopup(ShowPopup As Boolean)
  If Not (CurrentQuery.FieldByName("EVENTOANTERIOR").IsNull) Then
    Dim vHandle As Long
    ShowPopup = False
    vHandle = ProcuraGrausValido(GRAUANTERIOR.Text, "A")
    If (vHandle <> 0) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("GRAUANTERIOR").Value = vHandle
    End If
  End If
End Sub

Public Sub GRAUPOSTERIOR_OnPopup(ShowPopup As Boolean)
  If Not (CurrentQuery.FieldByName("EVENTOPOSTERIOR").IsNull) Then
    Dim vHandle As Long
    ShowPopup = False
    vHandle = ProcuraGrausValido(GRAUPOSTERIOR.Text, "P")

    If (vHandle <> 0) Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("GRAUPOSTERIOR").Value = vHandle
    End If
  End If
End Sub

Public Sub EVENTOANTERIOR_OnPopup(ShowPopup As Boolean)

 Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOANTERIOR.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOANTERIOR").Value = vHandle
  End If

End Sub

Public Sub EVENTOANULADOR_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOANULADOR.Text)
  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOANULADOR").Value = vHandle
  End If

End Sub

Public Sub GRAUANULADOR_OnPopup(ShowPopup As Boolean)

  If CurrentQuery.FieldByName("EVENTOANULADOR").IsNull Then
    bsShowMessage("Escolha primeiro o Evento Anulador", "I")
    ShowPopup = False
  Else
    Dim vHandle As Long
    ShowPopup = False
    vHandle = ProcuraGrausValido(GRAUANULADOR.Text, "X")
    If vHandle <> 0 Then
      CurrentQuery.Edit
      CurrentQuery.FieldByName("GRAUANULADOR").Value = vHandle
    End If
  End If

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
  vHandle = Interface.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE AND SAM_MOTIVOGLOSA.ATIVA='S']", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, MOTIVOGLOSAANTERIOR.Text)
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
  vHandle = Interface.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE AND SAM_MOTIVOGLOSA.ATIVA='S']", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, MOTIVOGLOSAPOSTERIOR.Text)
  Set Interface = Nothing

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").Value = vHandle
  End If
End Sub

Public Sub MULTIPLOSANTERIORES_OnClick()
  Dim Interface As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição","I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPOACAO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull Then
      bsShowMessage("Deve ser preenchido pelo menos um motivo de Glosa (Anterior/Posterior)","I")
      Exit Sub
    End If
  End If

  If VisibleMode Then 'As rotinas para desktop e web são feitas em lugar diferente
    Set Interface = CreateBennerObject("SAMDUPEVENTOS.ROTINAS")
    Set Interface.DuplicarIncomp(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "A")

  Else
    Dim viRetorno As Integer
    Dim vvContainer As CSDContainer
    Set vvContainer = NewContainer

    Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
        viRetorno = Interface.Exec(CurrentSystem, _
                                 1, _
                                 "TV_FORM0059", _
                                 "Eventos Anteriores - Incompatibilidade", _
                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                 150, _
                                 400, _
                                 False, _
                                 "", _
                                 vvContainer)
    Set vvContainer = Nothing
  End If
  RefreshNodesWithTable("SAM_INCOMP_EVENTOS_GERAL")
  Set Interface = Nothing
End Sub

Public Sub MULTIPLOSPOSTERIORES_OnClick()
  Dim Interface As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPOACAO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull Then
      bsShowMessage("Deve ser preenchido pelo menos um motivo de Glosa (Anterior/Posterior)", "I")
      Exit Sub
    End If
  End If

  If VisibleMode Then  'As rotinas para desktop e web são feitas em lugar diferente
    Set Interface = CreateBennerObject("SAMDUPEVENTOS.ROTINAS")
    Set Interface.DuplicarIncomp(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "P")

  Else
    Dim viRetorno As Integer
    Dim vvContainer As CSDContainer
    Set vvContainer = NewContainer

    Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
        viRetorno = Interface.Exec(CurrentSystem, _
                                 1, _
                                 "TV_FORM0060", _
                                 "Eventos Posteriores - Incompatibilidade", _
                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                 150, _
                                 400, _
                                 False, _
                                 "", _
                                 vvContainer)
    Set vvContainer = Nothing
  End If

  RefreshNodesWithTable("SAM_INCOMP_EVENTOS_GERAL")

  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  'sms 34639 - Edilson.Castro - 19/08/2005
  'o case foi reformulado para a nova estrutura das pastas
  If (WebMode) Then
    TIPO.ReadOnly = False
  Else
    TIPO.ReadOnly = True
  Select Case NodeInternalCode
    Case 62520050
      CurrentQuery.FieldByName("TIPO").AsString = "G"
      CurrentQuery.FieldByName("TIPOEVENTO").AsString = "M"
    Case 62520060
      CurrentQuery.FieldByName("TIPO").AsString = "G"
      CurrentQuery.FieldByName("TIPOEVENTO").AsString = "O"
    Case 62520150
      CurrentQuery.FieldByName("TIPO").AsString = "E"
      CurrentQuery.FieldByName("TIPOEVENTO").AsString = "M"
    Case 62520160
      CurrentQuery.FieldByName("TIPO").AsString = "E"
      CurrentQuery.FieldByName("TIPOEVENTO").AsString = "O"
  End Select
  End If

  Dim Valor As Long

  If EVENTOANTERIOR.ReadOnly Then
    CurrentQuery.FieldByName("EVENTOANTERIOR").AsInteger = RecordHandleOfTable("SAM_TGE")
  End If
  If EVENTOPOSTERIOR.ReadOnly Then
    CurrentQuery.FieldByName("EVENTOPOSTERIOR").AsInteger = RecordHandleOfTable("SAM_TGE")
  End If

  NewCounter("SAM_INCOMP_EVENTOS_GERAL", CurrentBranch, 1, Valor)
  CurrentQuery.FieldByName("INCOMPATIBILIDADE").AsInteger = Valor

End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    EVENTOPOSTERIOR.WebLocalWhere = "A.ULTIMONIVEL = 'S'"
    GRAUANTERIOR.WebLocalWhere = "A.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE WHERE GE.EVENTO=@CAMPO(EVENTOANTERIOR) " + _
                                 " AND GE.GRAU=A.HANDLE) )"
    GRAUPOSTERIOR.WebLocalWhere = "A.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE WHERE GE.EVENTO=@CAMPO(EVENTOPOSTERIOR) " + _
                                  " AND GE.GRAU=A.HANDLE) )"
    EVENTOANTERIOR.WebLocalWhere = "A.ULTIMONIVEL = 'S'"
    EVENTOANULADOR.WebLocalWhere = "A.ULTIMONIVEL = 'S'"
    GRAUANULADOR.WebLocalWhere = "(A.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE WHERE GE.EVENTO=@CAMPO(EVENTOANULADOR) " + _
                                 " AND GE.GRAU=A.HANDLE)))"
    MOTIVOGLOSAANTERIOR.WebLocalWhere = "A.ATIVA='S'"
    MOTIVOGLOSAPOSTERIOR.WebLocalWhere = "A.ATIVA='S'"
  End If

  TableReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  INATIVA.Visible = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As BPesquisa
  Set SQL = NewQuery

  If Not (VigenciaValida()) Then
	CanContinue = False
    bsShowMessage("Vigência inválida, data final não pode ser inferior a data inicial.", "E")
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("EVENTOPOSTERIOR").IsNull And CurrentQuery.FieldByName("EVENTOANTERIOR").IsNull) Then
	CanContinue = False
    bsShowMessage("É obrigatório informar ao menos um dos campos: 'Evento anterior' ou 'Evento posterior'.", "E")
    Exit Sub
  Else
	If (CurrentQuery.FieldByName("EVENTOPOSTERIOR").IsNull And CurrentQuery.FieldByName("GRAUPOSTERIOR").IsNull) Then
	  CanContinue = False
	  bsShowMessage("É obrigatório informar ao menos um dos campos: 'Evento posterior' ou 'Grau posterior'.", "E")
	  Exit Sub
	Else
	  If (CurrentQuery.FieldByName("EVENTOANTERIOR").IsNull And CurrentQuery.FieldByName("GRAUANTERIOR").IsNull) Then
	    CanContinue = False
	    bsShowMessage("É obrigatório informar ao menos um dos campos: 'Evento anterior' ou 'Grau anterior'.", "E")
		Exit Sub
	  End If
	End If
  End If

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT HANDLE")
  SQL.Add("  FROM SAM_INCOMP_EVENTOS_GERAL")
  SQL.Add(" WHERE HANDLE         <> " + CurrentQuery.FieldByName("HANDLE").AsString)

  If CurrentQuery.FieldByName("EVENTOANTERIOR").IsNull Then
    SQL.Add(" AND EVENTOANTERIOR IS NULL")
  Else
    SQL.Add(" AND EVENTOANTERIOR = " + CurrentQuery.FieldByName("EVENTOANTERIOR").AsString)
  End If

  If CurrentQuery.FieldByName("GRAUANTERIOR").IsNull Then
    SQL.Add(" AND GRAUANTERIOR IS NULL")
  Else
    SQL.Add(" AND GRAUANTERIOR = " + CurrentQuery.FieldByName("GRAUANTERIOR").AsString)
  End If

  If CurrentQuery.FieldByName("EVENTOPOSTERIOR").IsNull Then
    SQL.Add(" AND EVENTOPOSTERIOR IS NULL")
  Else
    SQL.Add(" AND EVENTOPOSTERIOR = " + CurrentQuery.FieldByName("EVENTOPOSTERIOR").AsString)
  End If

  If CurrentQuery.FieldByName("GRAUPOSTERIOR").IsNull Then
    SQL.Add(" AND GRAUPOSTERIOR IS NULL")
  Else
    SQL.Add(" AND GRAUPOSTERIOR = " + CurrentQuery.FieldByName("GRAUPOSTERIOR").AsString)
  End If

  SQL.Add("   AND CONSIDERAEXECUTOR = '" + CurrentQuery.FieldByName("CONSIDERAEXECUTOR").AsString + "'")
  SQL.Add("   AND TIPO              = '" + CurrentQuery.FieldByName("TIPO").AsString +"'") 'sms 55308 - Edilson.Castro - 22/12/2005

  'Soares - SMS: 60815 - 23/05/2006 - Início
  'Faz a verificacao se está marcado no tab como nao considerar, se estiver permite incluir outra incompatibilidade com mesmo eventos e graus.
  SQL.Add("   AND (TABTIPOACAO    <> 3 ")
  SQL.Add("    or TABTIPOACAONEGACAO <> 3) ")
  'Soares - SMS: 60815 - 23/05/2006 - Fim

  SQL.Active = True

  If (Not SQL.FieldByName("HANDLE").IsNull) And _
    ((CurrentQuery.FieldByName("TABTIPOACAO").AsInteger <> 3) Or (CurrentQuery.FieldByName("TABTIPOACAONEGACAO").AsInteger <> 3)) Then 'Soares - SMS: 60815 - 23/05/2006 - Início
    CanContinue = False
    bsShowMessage("Já existe uma incompatibilidade com os mesmos eventos e graus cadastrada no sistema", "E")

    Set SQL = Nothing
    Exit Sub
  End If

  'If TABTIPOVERIFICACAO.PageIndex = 0 Then 'Jun - SMS 40836 - 23/03/2005
    If CurrentQuery.FieldByName("TABTIPOACAO").AsInteger = 1 Then
      If CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull Then
        bsShowMessage("Deve ser preenchido pelo menos um motivo de Glosa (Anterior/Posterior)","E")
        CanContinue = False
        Exit Sub
      End If

      If(Not CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 0)Or _
         (Not CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull And CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 0)Then
        bsShowMessage("Quando selecionado um motivo de Glosa, deve-se preencher o campo % de Glosa !","E")
        CanContinue = False
        Exit Sub
    End If

  If(CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat <>0)And(CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").IsNull)Then
    CanContinue = False
    bsShowMessage("Motivo de glosa do evento anterior é obrigatório quando percentual da glosa é diferente de zero", "E")
  End If

  If(CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat <>0)And(CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull)Then
    CanContinue = False
    bsShowMessage("Motivo de glosa do evento posterior é obrigatório quando percentual da glosa é diferente de zero", "E")
  End If
End If
'Jun - SMS 40836 - 23/03/2005 - Inicio
'Else
  If CurrentQuery.FieldByName("TABTIPOACAONEGACAO").AsInteger = 1 Then
    If CurrentQuery.FieldByName("MOTIVONEGACAOANTERIOR").IsNull And CurrentQuery.FieldByName("MOTIVONEGACAOPOSTERIOR").IsNull Then
      bsShowMessage("Deve ser preenchido pelo menos um motivo de negação (Anterior/Posterior)","E")
      CanContinue = False
      Exit Sub
    End If
  End If
'End If 'If CurrentQuery.FieldByName("TABTIPOVERIFICACAO").AsInteger = 1 Then
'Jun - SMS 40836 - 23/03/2005 - Final

SQL.Clear
SQL.Add("SELECT G1.HANDLE FROM SAM_GRAU G1")
SQL.Add("WHERE G1.HANDLE IN (SELECT G2.HANDLE FROM SAM_GRAU G2 WHERE G2.VERIFICAGRAUSVALIDOS = 'N' OR")
SQL.Add("                           (EXISTS (SELECT GE.HANDLE FROM SAM_TGE_GRAU GE ")
SQL.Add("                            WHERE GE.EVENTO=:EVENTO AND GE.GRAU=G2.HANDLE)))")
SQL.Add("AND G1.HANDLE=:GRAU")

If (Not CurrentQuery.FieldByName("GRAUANTERIOR").IsNull And Not CurrentQuery.FieldByName("EVENTOANTERIOR").IsNull) Then
  SQL.Active = False
  SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTOANTERIOR").AsInteger
  SQL.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAUANTERIOR").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    CanContinue = False
    Set SQL = Nothing
    bsShowMessage("Grau Anterior não é válido para o Evento Anterior !", "E")
    Exit Sub
  End If
End If

If (Not CurrentQuery.FieldByName("GRAUPOSTERIOR").IsNull And Not CurrentQuery.FieldByName("EVENTOPOSTERIOR").IsNull) Then
  SQL.Active = False
  SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTOPOSTERIOR").AsInteger
  SQL.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAUPOSTERIOR").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    CanContinue = False
    Set SQL = Nothing
    bsShowMessage("Grau Posterior não é válido para o Evento Posterior !", "E")
    Exit Sub
  End If
End If

Dim validaCBOCBOS As Boolean
Dim dllValidarCBOCBOS As Object
Set dllValidarCBOCBOS = CreateBennerObject("Especifico.uEspecifico")
validaCBOCBOS = dllValidarCBOCBOS.PRO_ValidaConsideraCBOCBOS(CurrentSystem)

'SMS 167056 - Anderson Silva
If ((validaCBOCBOS And (CurrentQuery.FieldByName("CONSIDERACBOCBOS").AsString = "M")) Or (CurrentQuery.FieldByName("CONSIDERACBOCBOS").AsString = "C")) Then
  If Not ((CurrentQuery.FieldByName("CONSIDERAEXECUTOR").AsString = "M") And (CurrentQuery.FieldByName("CONSIDERALOCALEXECUCAO").AsString = "M")) Then
    bsShowMessage("CBO/CBOS – Parametrização permitida somente para “mesmos” Executores e Locais de Execução!", "E")
    CanContinue = False
    Exit Sub
  End If
End If
'SMS 167056 - Anderson Silva

If CurrentQuery.FieldByName("TABTIPOACAO").AsInteger >1 Then
  CurrentQuery.FieldByName("MOTIVOGLOSAANTERIOR").Clear
  CurrentQuery.FieldByName("MOTIVOGLOSAPOSTERIOR").Clear
  CurrentQuery.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 0
  CurrentQuery.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 0
  CurrentQuery.FieldByName("MOTIVONEGACAO").Clear
End If
Set SQL = Nothing

If CurrentQuery.FieldByName("TABPORFACE").AsInteger = 3 Then
  If Not(CurrentQuery.FieldByName("GRAUANTERIOR").IsNull And CurrentQuery.FieldByName("GRAUPOSTERIOR").IsNull) Then
    bsShowMessage("Campos grau anterior e grau posterior não devem ser informados quando a incompatibilidade for por face duplicada!", "E")
    CanContinue = False
    Exit Sub
  End If
  If CurrentQuery.FieldByName("PORDENTE").AsString = "S" Then
    bsShowMessage("Incompatibilidade por face duplicada não pode ser cadastrada junto com incompatibilidade por dente!", "E")
    CanContinue = False
    Exit Sub
  End If
End If

  If CurrentQuery.FieldByName("TIPOEVENTO").AsString <> "O" Then
     CurrentQuery.FieldByName("PORDENTE").AsString = "N"
     CurrentQuery.FieldByName("TIPORESTRICAOHISTORICO").AsString = "N"
     CurrentQuery.FieldByName("TIPORESTRICAOTRATAMENTO").AsString = "N"
  Else
     If CurrentQuery.FieldByName("TIPORESTRICAOHISTORICO").AsString = "N" And CurrentQuery.FieldByName("TIPORESTRICAOTRATAMENTO").AsString = "N" Then
       CurrentQuery.FieldByName("TIPORESTRICAOHISTORICO").AsString = "S"
       CurrentQuery.FieldByName("TIPORESTRICAOTRATAMENTO").AsString = "S"
       If CurrentQuery.State <> 3 Then
          bsShowMessage("A incompatibilidade odontológica exige ao menos um tipo de restrição!", "E")
          CanContinue = False
          Exit Sub
       End If
   End If
  End If

  If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    FinalizarVigenciasEspecificas
  End If

  Set dllValidarCBOCBOS = Nothing

End Sub

Function ProcuraGrausValido(descricaoDigitada As String, vTipo As String) As Long

  Dim Interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela As String
  Dim qParam As BPesquisa
  Set qParam = NewQuery

  qParam.Add("SELECT FILTRARGRAUSVALIDOS FROM SAM_PARAMETROSATENDIMENTO")
  qParam.Active = True

  Set Interface = CreateBennerObject("Procura.Procurar")

  Dim vHandle As Long
  vColunas = "DISTINCT SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_TIPOGRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"
  If qParam.FieldByName("FILTRARGRAUSVALIDOS").AsString = "S" Then
    If vTipo = "A" Then
      vCriterio = "(SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM     SAM_TGE_GRAU GE WHERE GE.EVENTO=" + CurrentQuery.FieldByName("EVENTOANTERIOR").AsString + " AND     GE.GRAU=SAM_GRAU.HANDLE)))"
    ElseIf vTipo = "P" Then
      vCriterio = "(SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM     SAM_TGE_GRAU GE WHERE GE.EVENTO=" + CurrentQuery.FieldByName("EVENTOPOSTERIOR").AsString + " AND     GE.GRAU=SAM_GRAU.HANDLE)))"
    Else
      vCriterio = "(SAM_GRAU.VERIFICAGRAUSVALIDOS = 'N' OR (EXISTS (SELECT GE.HANDLE FROM     SAM_TGE_GRAU GE WHERE GE.EVENTO=" + CurrentQuery.FieldByName("EVENTOANULADOR").AsString + " AND     GE.GRAU=SAM_GRAU.HANDLE)))"
    End If
  Else
    vCriterio = ""
  End If
  vCampos = "Código do Grau|Descrição|Tipo do Grau|Graus válidos"
  vTabela = "SAM_GRAU|SAM_TIPOGRAU[SAM_GRAU.TIPOGRAU = SAM_TIPOGRAU.HANDLE]"

  ProcuraGrausValido = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, "Graus de Atuação", True, descricaoDigitada)

  Set Interface = Nothing
  Set qParam = Nothing

End Function

Function VigenciaValida() As Boolean
	VigenciaValida = ((CurrentQuery.FieldByName("DATAFINAL").IsNull) Or (CurrentQuery.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAINICIAL").AsDateTime))
End Function

Private Sub FinalizarVigenciasEspecificas()
	Dim qUpdateIncompatibilidades As Object
	Set qUpdateIncompatibilidades = NewQuery

	If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
		Exit Sub
	End If

	qUpdateIncompatibilidades.Clear()
	qUpdateIncompatibilidades.Add("UPDATE SAM_INCOMP_EVENTOS_ESTADO SET DATAFINAL = :DATAFINAL WHERE INCOMPATIBILIDADE = :INCOMPATIBILIDADE AND DATAFINAL IS NULL")
	qUpdateIncompatibilidades.ParamByName("DATAFINAL").AsDateTime        = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	qUpdateIncompatibilidades.ParamByName("INCOMPATIBILIDADE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qUpdateIncompatibilidades.ExecSQL()

	qUpdateIncompatibilidades.Clear()
	qUpdateIncompatibilidades.Add("UPDATE SAM_INCOMP_EVENTOS_MUNICIPIO SET DATAFINAL = :DATAFINAL WHERE INCOMPATIBILIDADE = :INCOMPATIBILIDADE AND DATAFINAL IS NULL")
	qUpdateIncompatibilidades.ParamByName("DATAFINAL").AsDateTime        = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	qUpdateIncompatibilidades.ParamByName("INCOMPATIBILIDADE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qUpdateIncompatibilidades.ExecSQL()

	qUpdateIncompatibilidades.Clear()
	qUpdateIncompatibilidades.Add("UPDATE SAM_INCOMP_EVENTOS_PRESTADOR SET DATAFINAL = :DATAFINAL WHERE INCOMPATIBILIDADE = :INCOMPATIBILIDADE AND DATAFINAL IS NULL")
	qUpdateIncompatibilidades.ParamByName("DATAFINAL").AsDateTime        = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
	qUpdateIncompatibilidades.ParamByName("INCOMPATIBILIDADE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qUpdateIncompatibilidades.ExecSQL()

	CurrentQuery.FieldByName("INATIVA").AsString = "S"

	Set qUpdateIncompatibilidades = Nothing
End Sub
