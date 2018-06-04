'HASH: A5970E1013EFE22796769BA41443F841
Dim giEventoAnteriorPago As Long
Dim giEventoPosteriorPago As Long

'#Uses "*bsShowMessage
'#Uses "*CheckFilialProcessamento
'#Uses "*Arredonda
'#Uses "*UltimoDiaCompetencia

Public Sub TABLE_AfterScroll()
Dim qLote As Object
Dim qGuiaEventosAnterior As Object
Dim qGuiaEventosPosterior As Object
Dim qIncompGlosa As Object
Dim vvSamPegDLL As Object
Dim vvSamAcertos As Object
Dim q1 As Object
Dim qAux As Object
Dim vDatas As Date
Dim HandleIncompGlosa As Long

HandleIncompGlosa = RecordHandleOfTable("SAM_INCOMP_GLOSA")

Set qLote = NewQuery
Set qGuiaEventosAnterior = NewQuery
Set qGuiaEventosPosterior = NewQuery
Set qIncompGlosa = NewQuery
Set q1 = NewQuery
Set qAux = NewQuery

If CurrentQuery.State = 3 Then

  qIncompGlosa.Clear
  qIncompGlosa.Add("SELECT * FROM SAM_INCOMP_GLOSA ")
  qIncompGlosa.Add(" WHERE HANDLE = :HANDLE ")
  qIncompGlosa.ParamByName("HANDLE").AsInteger = HandleIncompGlosa
  qIncompGlosa.Active = True

  qLote.Clear
  qLote.Add("SELECT * FROM SAM_ACERTOLOTE ")
  qLote.Add(" WHERE SITUACAO = 'A' AND USUARIO = :USUARIO ")
  qLote.Add(" ORDER BY SEQUENCIA DESC, DATA DESC ")
  qLote.ParamByName("USUARIO").AsInteger = CurrentUser
  qLote.Active = True
  If (Not qLote.EOF) Then
    CurrentQuery.FieldByName("NUMEROLOTE").AsInteger = qLote.FieldByName("HANDLE").AsInteger
  End If
  qLote.Active = False

  qGuiaEventosAnterior.Clear
  qGuiaEventosAnterior.Add("SELECT * FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HGUIAEVENTO")
  qGuiaEventosAnterior.ParamByName("HGUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
  qGuiaEventosAnterior.Active = True

  qGuiaEventosPosterior.Clear
  qGuiaEventosPosterior.Add("SELECT * FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HGUIAEVENTO")
  qGuiaEventosPosterior.ParamByName("HGUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger
  qGuiaEventosPosterior.Active = True

  Set vvSamPegDLL = CreateBennerObject("SAMPEG.Rotinas")

  CurrentQuery.FieldByName("QUANTIDADEANT").AsFloat = Arredonda(qGuiaEventosAnterior.FieldByName("QTDPAGTO").AsFloat * qIncompGlosa.FieldByName("PERCENTGLOSAANTERIOR").AsFloat / 100)
  CurrentQuery.FieldByName("QUANTIDADEPOS").AsFloat = Arredonda(qGuiaEventosPosterior.FieldByName("QTDPAGTO").AsFloat * qIncompGlosa.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat / 100)

  CurrentQuery.FieldByName("VALORAGLOSARANT").AsFloat = Arredonda(vvSamPegDLL.SugerirValorGlosa(CurrentSystem, _
                                                                                                qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger, _
                                                                                                qGuiaEventosAnterior.FieldByName("QTDPAGTO").AsFloat) * qIncompGlosa.FieldByName("PERCENTGLOSAANTERIOR").AsFloat / 100)
  CurrentQuery.FieldByName("VALORAGLOSARPOS").AsFloat = Arredonda(vvSamPegDLL.SugerirValorGlosa(CurrentSystem, _
                                                                                                qIncompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger, _
                                                                                                qGuiaEventosPosterior.FieldByName("QTDPAGTO").AsFloat) * qIncompGlosa.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat / 100)


  Set vvSamAcertos = CreateBennerObject("SamAcertos.Incompatibilidade")
  CurrentQuery.FieldByName("PFADEVOLVERANT").AsFloat = vvSamAcertos.AtualizaPF(CurrentSystem, _
                                                                               qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger, _
                                                                               CurrentQuery.FieldByName("QUANTIDADEANT").AsFloat, _
                                                                               CurrentQuery.FieldByName("VALORAGLOSARANT").AsFloat)
  CurrentQuery.FieldByName("PFADEVOLVERPOS").AsFloat = vvSamAcertos.AtualizaPF(CurrentSystem, _
                                                                               qIncompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger, _
                                                                               CurrentQuery.FieldByName("QUANTIDADEPOS").AsFloat, _
                                                                               CurrentQuery.FieldByName("VALORAGLOSARPOS").AsFloat)
  Set vvSamPegDLL = Nothing
  Set vvSamAcertos = Nothing

  If QUANTIDADEANT.ReadOnly Then
    CurrentQuery.FieldByName("MOTIVOANT").Clear
  Else
    CurrentQuery.FieldByName("MOTIVOANT").AsInteger = qIncompGlosa.FieldByName("MOTIVOGLOSAANTERIOR").AsInteger
  End If

  If QUANTIDADEPOS.ReadOnly Then
    CurrentQuery.FieldByName("MOTIVOPOS").Clear
  Else
    CurrentQuery.FieldByName("MOTIVOPOS").AsInteger = qIncompGlosa.FieldByName("MOTIVOGLOSAPOSTERIOR").AsInteger
  End If

  'sugerir a data Base de irrf, de acordo com os parâmetros gerais fornecidos
  q1.Clear
  q1.Add("SELECT TABACERTOIRRF, DATABASEIRRF FROM SAM_PARAMETROSPROCCONTAS")
  q1.Active = True

  If q1.FieldByName("TABACERTOIRRF").AsInteger = 1 Then 'gerar irrf
    DATABASEIRRF.ReadOnly = False
    If q1.FieldByName("DATABASEIRRF").AsString = "O" Then 'original
      q1.Active = False
      q1.Clear
      q1.Add("SELECT F.COMPETENCIAIRRF FROM SFN_FATURA F, SAM_GUIA_EVENTOS E WHERE E.HANDLE=:GUIAEVENTO AND F.HANDLE=E.FATURAPAGAMENTO")
      q1.ParamByName("GUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
      q1.Active = True
      CurrentQuery.FieldByName("DATABASEIRRF").AsDateTime = q1.FieldByName("COMPETENCIAIRRF").AsDateTime
      q1.Active = False
    Else 'data Do acerto
      CurrentQuery.FieldByName("DATABASEIRRF").AsDateTime = ServerDate
    End If
  Else
    DATABASEIRRF.ReadOnly = True
  End If

  'INSS
  'sugerir a data Base de INSS, de acordo com os parâmetros gerais fornecidos
  q1.Active = False
  q1.Clear
  q1.Add("SELECT TABACERTOINSS, DATABASEINSS FROM SAM_PARAMETROSPROCCONTAS")
  q1.Active = True
  If q1.FieldByName("TABACERTOINSS").AsInteger = 1 Then 'gerar INSS
    COMPETINSS.ReadOnly = False
    If q1.FieldByName("DATABASEINSS").AsString = "O" Then 'original
      q1.Active = False
      q1.Clear
      q1.Add("SELECT F.COMPETENCIAINSS FROM SFN_FATURA F, SAM_GUIA_EVENTOS E WHERE E.HANDLE=:GUIAEVENTO AND F.HANDLE=E.FATURAPAGAMENTO")
      q1.ParamByName("GUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
      q1.Active = True
      CurrentQuery.FieldByName("COMPETINSS").AsDateTime = q1.FieldByName("COMPETENCIAINSS").AsDateTime
      q1.Active = False
    Else 'data Do acerto
      CurrentQuery.FieldByName("COMPETINSS").AsDateTime = ServerDate
    End If
  Else
    COMPETINSS.ReadOnly = True
  End If

  'ISS
  'sugerir a data Base de ISS, de acordo com os parâmetros gerais fornecidos
  q1.Active = False
  q1.Clear
  q1.Add("SELECT TABACERTOISS, DATABASEISS FROM SAM_PARAMETROSPROCCONTAS")
  q1.Active = True
  If q1.FieldByName("TABACERTOISS").AsInteger = 1 Then 'gerar ISS
    COMPETISS.ReadOnly = False
    If q1.FieldByName("DATABASEISS").AsString = "O" Then 'original
      q1.Active = False
      q1.Clear
      q1.Add("SELECT F.COMPETENCIAISS FROM SFN_FATURA F, SAM_GUIA_EVENTOS E WHERE E.HANDLE=:GUIAEVENTO AND F.HANDLE=E.FATURAPAGAMENTO")
      q1.ParamByName("GUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
      q1.Active = True
      CurrentQuery.FieldByName("COMPETISS").AsDateTime = q1.FieldByName("COMPETENCIAISS").AsDateTime
      q1.Active = False
    Else 'data Do acerto
      CurrentQuery.FieldByName("COMPETISS").AsDateTime = ServerDate
    End If
  Else
    COMPETISS.ReadOnly = False
  End If

  If (Not qGuiaEventosAnterior.FieldByName("BENEFICIARIO").IsNull) Then
    qAux.Active = False
    qAux.Clear
    qAux.Add(" SELECT X.DIACOBRANCA FROM (")
    qAux.Add(" SELECT F.DIACOBRANCA")
    qAux.Add("  FROM SAM_BENEFICIARIO B, ")
    qAux.Add("     SAM_FAMILIA F, ")
    qAux.Add("     SAM_CONTRATO C ")
    qAux.Add(" WHERE F.HANDLE = B.FAMILIA ")
    qAux.Add(" AND C.HANDLE = F.CONTRATO ")
    qAux.Add(" AND B.HANDLE = :BENEFICIARIO ")
    qAux.Add(" AND C.LOCALFATURAMENTO = 'F' ")
    qAux.Add(" UNION ALL")
    qAux.Add(" SELECT C.DIACOBRANCA")
    qAux.Add("  FROM SAM_BENEFICIARIO B, ")
    qAux.Add("     SAM_FAMILIA F, ")
    qAux.Add("     SAM_CONTRATO C ")
    qAux.Add(" WHERE F.HANDLE = B.FAMILIA ")
    qAux.Add(" AND C.HANDLE = F.CONTRATO ")
    qAux.Add(" AND B.HANDLE = :BENEFICIARIO ")
    qAux.Add(" AND C.LOCALFATURAMENTO <> 'F' ")
    qAux.Add(" ) X")

    qAux.ParamByName("BENEFICIARIO").AsInteger = qGuiaEventosAnterior.FieldByName("BENEFICIARIO").AsInteger
    qAux.Active = True

    If qAux.FieldByName("DIACOBRANCA").AsInteger <> 0 Then
      'decodedate(Sys.ServerDate, ano, mes, dia);
      If (qAux.FieldByName("DIACOBRANCA").AsInteger > Day(ServerDate)) Then
          If (Month(ServerDate) = 2 And qAux.FieldByName("DIACOBRANCA").AsInteger > 28) Then
            vDatas = UltimoDiaCompetencia(DateSerial(Year(ServerDate), _
                                                   Month(ServerDate), _
                                                   1))
          ElseIf ((Month(ServerDate) = 4 Or _
                   Month(ServerDate) = 6 Or _
                   Month(ServerDate) = 9 Or _
                   Month(ServerDate) = 11) _
                  And qAux.FieldByName("DIACOBRANCA").AsInteger > 30) Then
            vDatas = UltimoDiaCompetencia(DateSerial(Year(ServerDate), _
                                                   Month(ServerDate), _
                                                   1))
          Else
            vDatas = DateSerial(Year(ServerDate), _
                                Month(ServerDate), _
                                qAux.FieldByName("DIACOBRANCA").AsInteger)
          End If
      Else
        vDatas = DateAdd("M", _
                         1, _
                         DateSerial(Year(ServerDate), _
                                    Month(ServerDate), _
                                    qAux.FieldByName("DIACOBRANCA").AsInteger))

      End If
      CurrentQuery.FieldByName("DEBITOBENEFICIARIO").AsDateTime = vDatas
      CurrentQuery.FieldByName("CREDITOBENEFICIARIO").AsDateTime = vDatas
    End If
  End If
End If

VALORAGLOSARANT.ReadOnly = True
VALORAGLOSARPOS.ReadOnly = True
PFADEVOLVERANT.ReadOnly = True
PFADEVOLVERPOS.ReadOnly = True
MOTIVOANT.ReadOnly = True
MOTIVOPOS.ReadOnly = True

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

Dim Interface As Object
Dim vvDLLSamUtil As Object
Dim VsResultado As String
Dim HandleIncompGlosa As Long
Dim q1 As Object
Dim vbContinua As Boolean
Dim vsMensagem As String
Dim viRetorno As Long
Dim vsTipoMensagem As String
Dim qIncompGlosa As Object
Dim qSeleciona As Object
Dim viEventoAnteriorGlosado As Long
Dim viEventoPosteriorGlosado As Long

Set qIncompGlosa = NewQuery
Set qSeleciona = NewQuery
Set q1 = NewQuery

Set vvDLLSamUtil = CreateBennerObject("SAMUTIL.ROTINAS")

vvDLLSamUtil.CriaTabelaTemporariaSqlServer(CurrentSystem, 9)

Set vvDLLSamUtil = Nothing

HandleIncompGlosa = RecordHandleOfTable("SAM_INCOMP_GLOSA")

qIncompGlosa.Clear
qIncompGlosa.Add("SELECT EVENTOGUIAANTERIOR, EVENTOGUIAPOSTERIOR FROM SAM_INCOMP_GLOSA ")
qIncompGlosa.Add(" WHERE HANDLE = :HANDLE ")
qIncompGlosa.ParamByName("HANDLE").AsInteger = HandleIncompGlosa
qIncompGlosa.Active = True
'Primeiro verificar se já existe um Mov. acerto aberto para este evento
'se tiver, não pode criar outro
qSeleciona.Clear
qSeleciona.Add("SELECT HANDLE FROM SAM_GUIA_EVENTOS_ACERTO WHERE " + _
               "GUIAEVENTO=:GUIAEVENTO AND ACERTOREALIZADO='N' ")
qSeleciona.ParamByName("GUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAANTERIOR").AsInteger
qSeleciona.Active = True

If (Not qSeleciona.EOF) Then 'ainda tem acerto não realizado
  qSeleciona.Active = False

  vsMensagem = "O evento possui acerto não realizado." + Chr(13) + _
               "Volte a executar esta operação após os acertos terem sido faturados." + Chr(13) + _
               "Operação cancelada!"
  bsShowMessage(vsMensagem, "E")
  CanContinue = False
  Exit Sub
End If

qSeleciona.Active = False
qSeleciona.ParamByName("GUIAEVENTO").AsInteger = qIncompGlosa.FieldByName("EVENTOGUIAPOSTERIOR").AsInteger
qSeleciona.Active = True

If (Not qSeleciona.EOF) Then 'ainda tem acerto não realizado
  qSeleciona.Active = False

  vsMensagem = "O evento possui acerto não realizado." + Chr(13) + _
               "Volte a executar esta operação após os acertos terem sido faturados." + Chr(13) + _
               "Operação cancelada!"
  bsShowMessage(vsMensagem, "E")
  CanContinue = False
  Exit Sub
End If



Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "A")
  If aux = False Then
    If VisibleMode Then
      RefreshNodesWithTable("SAM_INCOMP_GLOSA")'Colocar tabela para refresh
      Exit Sub
    End If
  End If

  q1.Clear
  q1.Add("SELECT *")
  q1.Add("  FROM SAM_INCOMP_GLOSA ")
  q1.Add(" WHERE HANDLE = :HANDLE")
  q1.ParamByName("HANDLE").Value = HandleIncompGlosa
  q1.Active = True

  If(q1.FieldByName("SITUACAO").AsString <>"P")Then
    bsShowMessage("A incompatibilidade não está pendente", "I")
    CanContinue = False
    Exit Sub
  End If

'esta tabela possui 2 handles de guiaevento verificar para cada um,se está pago ou não
'se estiver pago,gerar movimento de acerto,senão glosar os eventos

'para o evento anterior e posterior
If(q1.FieldByName("EVENTOGUIAANTERIOR").IsNull And _
   q1.FieldByName("MOTIVOGLOSAANTERIOR").IsNull And _
   q1.FieldByName("EVENTOGUIAPOSTERIOR").IsNull And _
   q1.FieldByName("MOTIVOGLOSAPOSTERIOR").IsNull)Then
  bsShowMessage("Não há eventos a serem glosados", "E")
  CanContinue = False
  Exit Sub
Else
  If(q1.FieldByName("PERCENTGLOSAANTERIOR").AsFloat = 0)And _
     (q1.FieldByName("PERCENTGLOSAPOSTERIOR").AsFloat = 0)Then
    bsShowMessage("Não existe definição de Percentual de Glosa", "E")
    CanContinue = False
    Exit Sub
  End If
End If

Set Interface = CreateBennerObject("SamAcertos.Incompatibilidade")
Interface.Consistencia(CurrentSystem, HandleIncompGlosa, VsResultado)

viEventoAnteriorGlosado = 0
viEventoPosteriorGlosado  = 0

vbContinua = False

If (Not WebMode) Then
  If VsResultado <> "" Then
    If bsShowMessage(VsResultado + Chr(13) + "Deseja Continuar ?", "Q") = vbYes Then
      vbContinua = True
    End If
  Else
    vbContinua = True
  End If

  If vbContinua Then
    vsMensagem = Interface.VerificaPendenciasEventoAnterior(CurrentSystem, HandleIncompGlosa, vsTipoMensagem, viEventoAnteriorGlosado)
    If (vsMensagem <> "") Then
      If (vsTipoMensagem = "I") Then
        bsShowMessage(vsMensagem,"I")
      ElseIf (vsTipoMensagem = "Q") Then
        vbContinua = False
        If bsShowMessage(vsMensagem + ". Continuar assim mesmo ?", vsTipoMensagem) = vbYes Then
          vbContinua = True
        End If
      End If
    End If
  End If

  If vbContinua Then
    vsMensagem = Interface.VerificaPendenciasEventoPosterior(CurrentSystem, HandleIncompGlosa, vsTipoMensagem, viEventoPosteriorGlosado)
    If (vsMensagem <> "") Then
      If (vsTipoMensagem = "I") Then
        bsShowMessage(vsMensagem,"I")
      ElseIf (vsTipoMensagem = "Q") Then
        vbContinua = False
        If bsShowMessage(vsMensagem + ". Continuar assim mesmo ?", vsTipoMensagem) = vbYes Then
          vbContinua = True
        End If
      End If
    End If
  End If
Else
  If VsResultado <> "" Then
    bsShowMessage(VsResultado, "I")
  End If
  vsMensagem = Interface.VerificaPendenciasEventoAnterior(CurrentSystem, HandleIncompGlosa, vsTipoMensagem, viEventoAnteriorGlosado)
  If (vsMensagem <> "") Then
    bsShowMessage(vsMensagem, "I")
  End If
  vsMensagem = Interface.VerificaPendenciasEventoPosterior(CurrentSystem, HandleIncompGlosa, vsTipoMensagem, viEventoPosteriorGlosado)
  If (vsMensagem <> "") Then
    bsShowMessage(vsMensagem, "I")
  End If
  vbContinua = True
End If

giEventoAnteriorPago = 0
giEventoPosteriorPago = 0

If vbContinua Then
  vsMensagem = ""
  viRetorno = Interface.Glosar(CurrentSystem, HandleIncompGlosa, _
                               giEventoAnteriorPago, giEventoPosteriorPago, _
                               viEventoAnteriorGlosado, viEventoPosteriorGlosado, _
                               vsMensagem)

  If (viRetorno = 0) Then
    bsShowMessage(vsMensagem + "Impossível Continuar, Verifique os erros!", "E")
    vbContinua = False
  ElseIf (viRetorno = 1) And (Trim(vsMensagem) <> "") Then
    bsShowMessage(vsMensagem, "I")
    If (viEventoAnteriorGlosado = 1) And (Trim(vsMensagem) <> "") Then
      vbContinua = False
    End If
    If (viEventoPosteriorGlosado = 1) And (Trim(vsMensagem) <> "") Then
      vbContinua = False
    End If
    If (viEventoAnteriorGlosado = 1) And (viEventoPosteriorGlosado = 1) Then
      vbContinua = False
    End If
  End If

End If

If (viEventoAnteriorGlosado = 1) Then
  QUANTIDADEANT.ReadOnly = True
End If

If (viEventoPosteriorGlosado = 1) Then
  QUANTIDADEPOS.ReadOnly = True
End If

CanContinue = vbContinua

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim Interface As Object
Dim P_QuantidadeAnt As Double
Dim P_QuantidadePos As Double
Dim P_ValorGlosarAnt As Double
Dim P_ValorGlosarPos As Double
Dim P_PFADevolverAnt As Double
Dim P_PFADevolverPos As Double
Dim P_MotivoAnt As Long
Dim P_MotivoPos As Long
Dim P_MotivoMvtoAcerto As Long
Dim P_DebitoRecebedor As Date
Dim P_CreditoRecebdor As Date
Dim P_DebitoBeneficiario As Date
Dim P_CreditoBeneficiario As Date
Dim P_DataBaseIRRF As Date
Dim P_CompetINSS As Date
Dim P_CompetISS As Date
Dim P_NumeroLote As Long
Dim HandleIncompGlosa As Long

Dim vsMensagem As String
Dim viResult As Long

P_QuantidadeAnt = CurrentQuery.FieldByName("QUANTIDADEANT").AsFloat
P_QuantidadePos = CurrentQuery.FieldByName("QUANTIDADEPOS").AsFloat

P_ValorGlosarAnt = CurrentQuery.FieldByName("VALORAGLOSARANT").AsFloat
P_ValorGlosarPos = CurrentQuery.FieldByName("VALORAGLOSARPOS").AsFloat

P_PFADevolverAnt = CurrentQuery.FieldByName("PFADEVOLVERANT").AsFloat
P_PFADevolverPos = CurrentQuery.FieldByName("PFADEVOLVERPOS").AsFloat

If (Not CurrentQuery.FieldByName("MOTIVOANT").IsNull) Then
  P_MotivoAnt = CurrentQuery.FieldByName("MOTIVOANT").AsInteger
Else
  P_MotivoAnt = 0
End If

If (Not CurrentQuery.FieldByName("MOTIVOPOS").IsNull) Then
  P_MotivoPos = CurrentQuery.FieldByName("MOTIVOPOS").AsInteger
Else
  P_MotivoPos = 0
End If

P_MotivoMvtoAcerto = CurrentQuery.FieldByName("MOTIVOMOVTOACERTO").AsInteger
P_DebitoRecebedor = CurrentQuery.FieldByName("DEBITORECEBEDOR").AsDateTime
P_CreditoRecebdor = CurrentQuery.FieldByName("CREDITORECEBEDOR").AsDateTime
P_DebitoBeneficiario = CurrentQuery.FieldByName("DEBITOBENEFICIARIO").AsDateTime
P_CreditoBeneficiario = CurrentQuery.FieldByName("CREDITOBENEFICIARIO").AsDateTime
P_DataBaseIRRF = CurrentQuery.FieldByName("DATABASEIRRF").AsDateTime
P_CompetINSS = CurrentQuery.FieldByName("COMPETINSS").AsDateTime
P_CompetISS = CurrentQuery.FieldByName("COMPETISS").AsDateTime
P_NumeroLote = CurrentQuery.FieldByName("NUMEROLOTE").AsInteger

HandleIncompGlosa = RecordHandleOfTable("SAM_INCOMP_GLOSA")

Set Interface = CreateBennerObject("SamAcertos.Incompatibilidade")
viResult = Interface.GerarMovimentoAcerto(CurrentSystem, _
                                          HandleIncompGlosa, _
                                          P_QuantidadeAnt, _
							              P_QuantidadePos, _
							              P_ValorGlosarAnt, _
           							      P_ValorGlosarPos, _
							              P_PFADevolverAnt, _
							              P_PFADevolverPos, _
							              P_MotivoAnt, _
							              P_MotivoPos, _
							              P_MotivoMvtoAcerto, _
							              P_DebitoRecebedor, _
							              P_CreditoRecebdor, _
							              P_DebitoBeneficiario, _
							              P_CreditoBeneficiario, _
							              P_DataBaseIRRF, _
							              P_CompetINSS, _
							              P_CompetISS, _
							              P_NumeroLote, _
							              giEventoAnteriorPago, _
							              giEventoPosteriorPago, _
							              vsMensagem)
If (viResult = 0) Then
  bsShowMessage(vsMensagem, "E")
  CanContinue = False
End If

End Sub
