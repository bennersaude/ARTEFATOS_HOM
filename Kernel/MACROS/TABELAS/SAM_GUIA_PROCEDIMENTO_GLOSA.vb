'HASH: 315094C1F0D234EA5FFA52AF2DA52F1E
'Macro: SAM_GUIA_EVENTOS_GLOSA
' SAM_GUIA_EVENTOS_GLOSA
' 05-09-2000 -larini


Option Explicit
'#Uses "*ProcuraEvento"
'#Uses "*Arredonda"
'#Uses "*IsInt"

Dim vExigeComplemento As String
Dim alterouRevisada As Boolean
Dim auxGuiaEvento As Long
Dim Reconsideracao As Boolean
Dim vSituacaoAnteriorEvento As String


Public Function VERIFICAINCOMP As Boolean
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Add("SELECT SITUACAO FROM SAM_PEG WHERE HANDLE=:PEG")
  q1.ParamByName("PEG").Value = RecordHandleOfTable("SAM_PEG")
  q1.Active = True
  VERIFICAINCOMP = True
  If q1.FieldByName("SITUACAO").AsString = "3" Then 'PRONTO
    Dim INTERFACE As Object
    Set INTERFACE = CreateBennerObject("SAMINCOMP.CHECK")
    INTERFACE.INICIALIZAR(CurrentSystem)
    VERIFICAINCOMP = INTERFACE.EXCLUIRINCOMPATIBILIDADES(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger)
    INTERFACE.FINALIZAR
    Set INTERFACE = Nothing
  End If
  Set q1 = Nothing
End Function

Public Sub BOTAOGLOSA_OnClick()

  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "A")
  If aux = False Then
    RefreshNodesWithTable("SAM_GUIA_EVENTOS_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("samPEG.PROCESSAR")
  'INTERFACE.INICIALIZAR(CurrentSystem)
  INTERFACE.DESCRICAOGLOSA(CurrentSystem, CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger)
  'INTERFACE.FINALIZAR
  Set INTERFACE = Nothing
End Sub

Public Sub BOTAORECONSIDERACAO_OnClick()

  If CurrentQuery.State <>1 Then
    MsgBox("O registro não pode estar em edição")
    Exit Sub
  End If

  Dim SQL1 As Object
  Set SQL1 = NewQuery

  If Not InTransaction Then StartTransaction

  SQL1.Active = False
  SQL1.Clear
  SQL1.Add("UPDATE SAM_GUIA_EVENTOS_GLOSA                     ")
  SQL1.Add(" SET VALORGLOSA = :VALORGLOSA,                    ")
  SQL1.Add("     QTDGLOSA   = :QTDGLOSA,                      ")
  SQL1.Add("     QTDRECONSIDERADA   = :QTDRECONSIDERADA,      ")
  SQL1.Add("     VALORRECONSIDERADO   = :VALORRECONSIDERADO   ")
  SQL1.Add(" WHERE HANDLE = :HANDLE                           ")
  SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL1.ParamByName("VALORGLOSA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("VALORGLOSA").AsFloat, "##0.00"))
  SQL1.ParamByName("QTDGLOSA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("QTDGLOSA").AsFloat, "##0.00"))
  SQL1.ParamByName("QTDRECONSIDERADA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat, "##0.00"))
  SQL1.ParamByName("VALORRECONSIDERADO").AsFloat = CDbl(Format(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat, "##0.00"))
  SQL1.ExecSQL

  If InTransaction Then Commit

  Set SQL1 = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True


  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT FILIALPROCESSAMENTO FROM SAM_GUIA_EVENTOS WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  SQL.Active = True

  Dim aux As Boolean
  'Somente processará se a filial que esta digitando tiver como filialprocessamnto a sua filial
  aux = CheckFilialProcessamento(CurrentSystem, SQL.FieldByName("FILIALPROCESSAMENTO").AsInteger, "A")

  'início trecho alterado por sms 68329 - Edilson.Castro - 19/09/2006
  Dim MsgAux As String

  MsgAux = VerificaReconsideracao
  If Len(MsgAux) > 0 Then
    MsgBox(MsgAux)
    Exit Sub
  End If

  ' - trecho abaixo movido para a função VerificaReconsideracao (sms 68329)

  'Jun - SMS 44807 - 01/06/2005 - Início
  'Verificar se o motivo de glosa está inserido na carga abaixo do usuário, se não estiver então nào será possível reconsiderar a glosa

  'Dim qUsuarioGlosa As Object
  'Set qUsuarioGlosa = NewQuery

  '(...)

  'Else
  '  Set qUsuarioGlosa = Nothing
  'End If

  'Jun - SMS 44807 - 01/06/2005 - Final
  'fim trecho alterado por sms 68329

  Reconsideracao = True

  If aux = False Then
    RefreshNodesWithTable("SAM_GUIA_EVENTOS_GLOSA")'Colocar tabela para refresh
    Set SQL = Nothing
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("SELECT COUNT(GE.HANDLE) TOTAL")
  SQL.Add("  FROM SAM_GUIA_EVENTOS GE, ")
  SQL.Add("       SAM_GUIA_EVENTOS_ACERTO GEA")
  SQL.Add(" WHERE GEA.GUIAEVENTO=GE.HANDLE")
  SQL.Add("   AND GEA.TABTIPOACERTO=1 And GE.HANDLE=" + Str(CurrentQuery.FieldByName("GUIAEVENTO").AsInteger))
  SQL.Active = True

  If SQL.FieldByName("TOTAL").AsInteger <>0 Then
    MsgBox("Reconsideração não permitida: Evento sofreu movimento de acerto do tipo alteração!")
    Set SQL = Nothing
    Exit Sub
  End If
  Set SQL = Nothing

  Dim SQL2 As Object
  Set SQL2 = NewQuery
  SQL2.Add("SELECT COUNT(GE.HANDLE) TOTAL")
  SQL2.Add("  FROM SAM_GUIA_EVENTOS GE, ")
  SQL2.Add("       SAM_GUIA_EVENTOS_ACERTO GEA")
  SQL2.Add(" WHERE GEA.GUIAEVENTO=GE.HANDLE")
  SQL2.Add("   AND GEA.ACERTOREALIZADO='N' And GE.HANDLE=" + Str(CurrentQuery.FieldByName("GUIAEVENTO").AsInteger))
  SQL2.Active = True

  If SQL2.FieldByName("TOTAL").AsInteger <>0 Then
    MsgBox("Reconsideração não permitida: Existe movimento de acerto não faturado para este Evento!")
    Set SQL2 = Nothing
    Exit Sub
  End If
  Set SQL2 = Nothing

  'início trecho alterado por sms 68329 - Edilson.Castro - 19/09/2006
  ' - trecho movido para a função VerificaReconsideracao
  'Dim q1 As Object
  'Set q1 = NewQuery

  '(...)
  '  Exit Sub
  'End If
  'Set q1 = Nothing
  'fim trecho alterado por sms 68329


  Dim Q As Object
  Set Q = NewQuery

  Q.Clear
  Q.Add("SELECT HANDLE,FATURAPAGAMENTO FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HGUIAEVENTO")
  Q.ParamByName("HGUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
  Q.Active = True
  If Q.FieldByName("FATURAPAGAMENTO").IsNull Then
    '  On Error GoTo Sair

    CurrentQuery.Edit

    If CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "S" Then
      QTDGLOSA.ReadOnly = False 'antes: "true" - sms 68329 - Edilson.Castro - 19/09/2006
      VALORGLOSA.ReadOnly = True
      'QTDRECONSIDERADA.ReadOnly = False
      'VALORRECONSIDERADO.ReadOnly = True
    Else
      QTDGLOSA.ReadOnly = True
      VALORGLOSA.ReadOnly = False 'antes: "true" - sms 68329 - Edilson.Castro - 19/09/2006
      'VALORRECONSIDERADO.ReadOnly = False
      'QTDRECONSIDERADA.ReadOnly = True
    End If

    'COMPLEMENTORECONSIDERACAO.ReadOnly = False
    'MOTIVORECONSIDERACAO.ReadOnly = False
    'VALORRECONSIDERADO.ReadOnly        = False
    'QTDRECONSIDERADA.ReadOnly          = False

    CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = CurrentQuery.FieldByName("VALORGLOSA").AsFloat
    CurrentQuery.FieldByName("QTDRECONSIDERADA").Value = CurrentQuery.FieldByName("QTDGLOSA").AsFloat

    alterouRevisada = True
    CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S"

    'QTDRECONSIDERADA.SetFocus
Sair :
    Exit Sub
  End If

  Set Q = Nothing

End Sub



Public Sub GLOSAREVISADA_OnChange()
  alterouRevisada = True
End Sub

Public Sub MOTIVOGLOSA_OnExit()
  If CurrentQuery.State <>1 Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Clear
    SQL.Add("SELECT GLOSADEPENDENTE FROM SAM_MOTIVOGLOSA WHERE HANDLE=:MOTIVOGLOSA")
    SQL.ParamByName("MOTIVOGLOSA").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active = True
    CurrentQuery.FieldByName("GLOSADEPENDENTE").Value = SQL.FieldByName("GLOSADEPENDENTE").AsString
    Set SQL = Nothing
  End If

End Sub

Public Sub MOTIVOGLOSA_OnPopup(ShowPopup As Boolean)

  'SMS 48762 - Correcao conforme feita no cliente
  Dim SQL As Object
  Set SQL = NewQuery
  On Error GoTo prox
  SQL.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA WHERE CODIGOGLOSA=:CODIGOGLOSA")
  SQL.ParamByName("CODIGOGLOSA").AsString = MOTIVOGLOSA.Text
  SQL.Active = True
  If Not SQL.EOF Then
    CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger = SQL.FieldByName("HANDLE").AsInteger
    ShowPopup = False
    HabilitaCamposGlosa
    Set SQL = Nothing
    Exit Sub
  End If
  Set SQL = Nothing
prox:

  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim interface As Object
  'Dim vCampo As String
  'Dim vColuna As Integer
  ShowPopup = False


  'vCampo = MOTIVOGLOSA.Text

  'If IsInt(vCampo) Then
  '  vColuna = 1
  'Else
  ' vColuna = 2
  'End If

  vColunas = "CODIGOGLOSA|SAM_MOTIVOGLOSA.DESCRICAO|SAM_TIPOMOTIVOGLOSA.DESCRICAO"
  'vColunas ="CODIGOGLOSA|SAM_MOTIVOGLOSA.Z_DESCRICAO|SAM_TIPOMOTIVOGLOSA.DESCRICAO"
  vCampos = "Código|Descrição|Tipo Glosa"
  vCriterio = "SAM_MOTIVOGLOSA.ATIVA='S'"
  ' SMS 41587 - PARA FICAR IGUAL AO MG.
  vCriterio = vCriterio + "AND SAM_MOTIVOGLOSA.HANDLE NOT IN ("
  vCriterio = vCriterio + "SELECT SAM_MOTIVOGLOSA_SIS.SAMMOTIVOGLOSA FROM SAM_MOTIVOGLOSA_SIS)"

  Set interface = CreateBennerObject("Procura.Procurar")
  vHandle = interface.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE]", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, "")
  'vHandle =interface.Exec(CurrentSystem,"SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE]",vColunas,vColuna,vCampos,vCriterio,"Tabela de Motivo de Glosas",True,vCampo)
  Set interface = Nothing

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("MOTIVOGLOSA").Value = vHandle
    HabilitaCamposGlosa
  End If

  'ShowPopup =False
End Sub

Public Sub QTDGLOSA_OnExit()
  Dim vReadOnlyOriginal As Boolean
  vReadOnlyOriginal = VALORGLOSA.ReadOnly
  'sms 41587
  If (CurrentQuery.State = 3 Or CurrentQuery.State = 2) Then
    If CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0 Then 'Alterado de AsInteger para AsFloat na SMS 67809 - 02.09.2006
      VALORGLOSA.ReadOnly = True
      'QTDRECONSIDERADA.ReadOnly = False
      'VALORRECONSIDERADO.ReadOnly = True

    Else
      VALORGLOSA.ReadOnly = False
      'VALORRECONSIDERADO.ReadOnly = False
      If vReadOnlyOriginal Then
        VALORGLOSA.SetFocus
      End If
    End If
  End If

  'If CurrentQuery.State =3 And Not CurrentQuery.FieldByName("QTDGLOSA").IsNull Then 'inserção
  If (VALORGLOSA.ReadOnly) And (CurrentQuery.State = 3 Or CurrentQuery.State = 2) Then

    'Luciano T. Alberti - SMS 56345 - 20/01/2006 - Inicio
    Dim q1 As Object
    Set q1 = NewQuery

    q1.Clear
    q1.Add("SELECT QTDAPRESENTADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=" + CurrentQuery.FieldByName("GUIAEVENTO").AsString)
    q1.Active = True
    If Arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat) > Arredonda(q1.FieldByName("QTDAPRESENTADA").AsFloat) Then
       CurrentQuery.FieldByName("QTDGLOSA").AsFloat = Arredonda(q1.FieldByName("QTDAPRESENTADA").AsFloat)
    End If
    'Luciano T. Alberti - SMS 56345 - 20/01/2006 - Fim

    Dim interface As Object
    Set interface = CreateBennerObject("sampeg.processar")
    CurrentQuery.FieldByName("VALORGLOSA").Value = interface.SugerirValorGlosa(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, CurrentQuery.FieldByName("QTDGLOSA").AsFloat)
    Set interface = Nothing
  End If

End Sub

Public Sub QTDRECONSIDERADA_OnExit()
  Dim vQtdeReconsiderada As Double
  Dim vQtdeGlosa         As Double

  Dim q1 As Object
  Set q1 = NewQuery


  q1.Clear
  q1.Add("SELECT VALORCALCPAGTO, QTDAPRESENTADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=" + CurrentQuery.FieldByName("GUIAEVENTO").AsString)
  q1.Active = True

  vQtdeReconsiderada  = arredonda(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat)
  vQtdeGlosa          = arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat)

  'Luciano T. Alberti - SMS 56345 - 20/01/2006
  If((CurrentQuery.State = 2 Or CurrentQuery.State = 3 Or CurrentQuery.State = 2 Or CurrentQuery.State = 2) And _
     arredonda(vQtdeReconsiderada) > 0 And arredonda(vQtdeGlosa) > 0) Then

    Dim interface  As Object
    Set interface=CreateBennerObject("sampeg.processar")
    CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = interface.SugerirValorReconsiderado(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat)
    Set interface=Nothing

  'CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = arredonda((CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat * CurrentQuery.FieldByName("VALORGLOSA").AsFloat) / CurrentQuery.FieldByName("QTDGLOSA").AsFloat)
  '  CurrentQuery.FieldByName("VALORRECONSIDERADO").Value =arredonda((CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat * Q1.FieldByName("VALORCALCPAGTO").AsFloat)/CurrentQuery.FieldByName("QTDGLOSA").AsFloat)
  Else
    If (CurrentQuery.State=2 And _
        arredonda(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat) > 0 And _
        arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat) = 0) Then
      CurrentQuery.FieldByName("QTDRECONSIDERADA").Value = CurrentQuery.FieldByName("QTDGLOSA").Value
    End If
  End If

  If (CurrentQuery.State = 3 Or CurrentQuery.State = 2) Then
    If (CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0) And _
       (CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0) Then
     CurrentQuery.FieldByName("VALORRECONSIDERADO").Value = 0
    End If
  End If


q1.Active = False
End Sub



Public Sub TABLE_AfterDelete()
  Dim interface As Object
  Set interface = CreateBennerObject("BSPro000.Rotinas")
  interface.RevisarEvento(CurrentSystem, auxGuiaEvento, "REVISAR", True)
  Set interface = Nothing
End Sub

Public Sub TABLE_AfterInsert()
  'COMPLEMENTORECONSIDERACAO.ReadOnly = False
  'MOTIVORECONSIDERACAO.ReadOnly = False
  'VALORRECONSIDERADO.ReadOnly = False
  'QTDRECONSIDERADA.ReadOnly = False
  If CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger = 1 Then
  'início sms 68329 - Edilson.Castro - 20/09/2006
  ' - só vai se saber que campos habilitar após selecionado o tipo de glosa (de valor ou de qtd)
    'QTDRECONSIDERADA.ReadOnly = True
    'VALORRECONSIDERADO.ReadOnly = True
    QTDGLOSA.ReadOnly = True
    VALORGLOSA.ReadOnly = True
  'fim sms 68329
  End If
End Sub

Public Sub TABLE_AfterCommitted()
  Dim interface As Object
  Set interface=CreateBennerObject("BSPro000.Rotinas")
  'Set interface = CreateBennerObject("sampeg.processar")
  interface.RevisarEvento(CurrentSystem, CurrentQuery.FieldByName("GUIAEVENTO").AsInteger, "REVISAR", True)
  Dim qglosasOK As Object
  Set qglosasOK = NewQuery
  qglosasOK.Clear
  qglosasOK.Add("SELECT COUNT(1) QTDE           ")
  qglosasOK.Add("  FROM SAM_GUIA_EVENTOS_GLOSA  ")
  qglosasOK.Add(" WHERE GUIAEVENTO = :GUIAEVENTO")
  qglosasOK.Add("  AND GLOSAREVISADA = 'N' ")
  qglosasOK.ParamByName("GUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
  qglosasOK.Active = True
  If qglosasOK.FieldByName("QTDE").AsInteger = 0 Then
    Dim qnegacoesOK As Object
    Set qnegacoesOK = NewQuery
    qnegacoesOK.Clear
    qnegacoesOK.Add("SELECT COUNT(1) QTDE           ")
    qnegacoesOK.Add("  FROM SAM_GUIA_EVENTOS_NEGACAO")
    qnegacoesOK.Add(" WHERE GUIAEVENTO = :GUIAEVENTO")
    qnegacoesOK.Add("  AND NEGACAOREVISADA = 'N' ")
    qnegacoesOK.ParamByName("GUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
    qnegacoesOK.Active = True
    If qnegacoesOK.FieldByName("QTDE").AsInteger = 0 Then
      Dim qqtdeventosok As Object
      Set qqtdeventosok = NewQuery
      Dim vHandleGuia As Long
      qqtdeventosok.Clear
      qqtdeventosok.Add("SELECT GUIA                                 ")
      qqtdeventosok.Add("  FROM SAM_GUIA_EVENTOS E                   ")
      qqtdeventosok.Add(" WHERE E.HANDLE = :GUIAEVENTO               ")
      qqtdeventosok.ParamByName("GUIAEVENTO").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
      qqtdeventosok.Active = True
      vHandleGuia = qqtdeventosok.FieldByName("GUIA").Value

      qqtdeventosok.Clear
      qqtdeventosok.Add("SELECT COUNT(E.HANDLE) QTDE                 ")
      qqtdeventosok.Add("  FROM SAM_GUIA         G,                  ")
      qqtdeventosok.Add("       SAM_GUIA_EVENTOS E                   ")
      qqtdeventosok.Add(" WHERE G.HANDLE = :GUIA                     ")
      qqtdeventosok.Add("   AND E.GUIA = G.HANDLE                    ")
      qqtdeventosok.Add("   AND E.SITUACAO = :SITUACAO               ")
      qqtdeventosok.ParamByName("GUIA").Value = vHandleGuia
      qqtdeventosok.ParamByName("SITUACAO").AsString = vSituacaoAnteriorEvento
      qqtdeventosok.Active = True
      If qqtdeventosok.FieldByName("QTDE").AsInteger = 0 Then
        RefreshNodesWithTable("SAM_GUIA")
      Else
        RefreshNodesWithTable("SAM_GUIA_EVENTOS")
      End If
      Set qqtdeventosok = Nothing
    End If
    Set qnegacoesOK = Nothing
  End If
  Set qglosasOK = Nothing
  Set interface = Nothing
End Sub


Public Sub TABLE_AfterScroll()
  'COMPLEMENTORECONSIDERACAO.ReadOnly = True
  'MOTIVORECONSIDERACAO.ReadOnly = True
  'VALORRECONSIDERADO.ReadOnly = True
  'QTDRECONSIDERADA.ReadOnly = True
  QTDGLOSA.ReadOnly = True
  VALORGLOSA.ReadOnly = True
  alterouRevisada = False

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT DECISAO, APLICARECURSO FROM SAM_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  SQL.Active = True
 ' If SQL.FieldByName("DECISAO").AsString = "A" Then
 '   ROTULODECISAO.Text = "Decisão: Administrativo"
 ' ElseIf SQL.FieldByName("DECISAO").AsString = "M" Then
 '   ROTULODECISAO.Text = "Decisão: Médico"
 ' ElseIf SQL.FieldByName("DECISAO").AsString = "O" Then
 '   ROTULODECISAO.Text = "Decisão: Odontológico"
 ' Else
 '   ROTULODECISAO.Text = "Decisão: ------- "
 ' End If

 ' If SQL.FieldByName("APLICARECURSO").AsString = "S" Then
 '   ROTULORECURSO.Text = "Aplica Recurso"
 ' Else
 '   ROTULORECURSO.Text = "Não Aplica Recurso"
 ' End If
  Dim qSituacaoEvento As Object
  Set qSituacaoEvento = NewQuery
  If CurrentQuery.FieldByName("GUIAEVENTO").AsInteger > 0 Then
    qSituacaoEvento.Clear
    qSituacaoEvento.Add("SELECT SITUACAO FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HANDLE")
    qSituacaoEvento.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("GUIAEVENTO").Value
    qSituacaoEvento.Active = True
    vSituacaoAnteriorEvento = qSituacaoEvento.FieldByName("SITUACAO").AsString
  Else
    qSituacaoEvento.Clear
    qSituacaoEvento.Add("SELECT SITUACAO FROM SAM_GUIA_EVENTOS WHERE HANDLE = :HANDLE")
    qSituacaoEvento.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_GUIA_EVENTOS")
    qSituacaoEvento.Active = True
    vSituacaoAnteriorEvento = qSituacaoEvento.FieldByName("SITUACAO").AsString
  End If
  Set qSituacaoEvento = Nothing
  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeCancel(CanContinue As Boolean)
 ' If MOTIVORECONSIDERACAO.ReadOnly = False Then
 '   COMPLEMENTORECONSIDERACAO.ReadOnly = True
 '   MOTIVORECONSIDERACAO.ReadOnly = True
 '   VALORRECONSIDERADO.ReadOnly = True
 '   QTDRECONSIDERADA.ReadOnly = True
 '   QTDGLOSA.ReadOnly = True
 '   VALORGLOSA.ReadOnly = True
 ' End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  CanContinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "E")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTable("SAM_GUIA_EVENTOS_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  '  If VERIFICAINCOMP=False Then
  '    CanContinue =False
  '    MsgBox("Existe incompatibilidade glosada para este evento")
  '    Exit Sub
  '  End If
  VERIFICAJAPAGO(CanContinue)
  If CanContinue = False Then
    MsgBox("Alteração não permitida nesta fase")
  Else
    If CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger <>1 Then
      MsgBox("Só é permitida a exclusão de glosas manuais")
      CanContinue = False
      auxGuiaEvento = 0
    Else
      auxGuiaEvento = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
    End If
  End If
End Sub


Public Sub VERIFICAJAPAGO(CanContinue As Boolean)
  Dim SITUACAO As String
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Clear
  q1.Add("SELECT G.SITUACAO FROM SAM_GUIA G WHERE G.HANDLE=:GUIA")
  q1.ParamByName("GUIA").Value = RecordHandleOfTable("SAM_GUIA")
  q1.Active = True
  SITUACAO = q1.FieldByName("SITUACAO").AsString
  q1.Active = False
  Set q1 = Nothing
  If SITUACAO = "4" Then
    CanContinue = False
  Else
    CanContinue = True
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)


  'CanContinue =CheckFilialProcessamento(CurrentSystem,RecordHandleOfTable("FILIAIS"),"A")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTable("SAM_GUIA_EVENTOS_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  'início sms 68329 - Edilson.Castro - 19/09/2006
  ' - antes estava liberando todos os campos; deve liberar apenas o campo correspondente ao tipo da glosa (qtd ou valor)
  '   mas isso só é possível depois de definido qual é a glosa; por isso, os campos só são liberados no onpopup do motivoglosa

  'If CurrentQuery.FieldByName("ORIGEMGLOSA").AsString = "1" Then
    'QTDRECONSIDERADA.ReadOnly = False
    'VALORRECONSIDERADO.ReadOnly = False
    '(...)

  'fim sms 68329

  '  If VERIFICAINCOMP=False Then
  '    CanContinue =False
  '    MsgBox("Existe incompatibilidade glosada para este evento")
  '    Exit Sub
  '  End If
  VERIFICAJAPAGO(CanContinue)
  If CanContinue = False Then
    MsgBox("Alteração não permitida nesta fase")
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  CanContinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "I")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTable("SAM_GUIA_EVENTOS_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If

  '  If VERIFICAINCOMP=False Then
  '     MsgBox("Existe incompatibilidade glosada para este evento")
  '     Exit Sub
  '  End If
  VERIFICAJAPAGO(CanContinue)
  If CanContinue = False Then
    MsgBox("Inclusão não permitida nesta fase")
    CanContinue = False
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim valorapresentado As Variant
  Dim valorcalculado As Variant

  QTDRECONSIDERADA_OnExit

  'SMS 41587
  If CurrentQuery.FieldByName("QTDGLOSA").AsFloat > 0 Then
    CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "S"
  Else
    CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "N"
  End If


  If(CurrentQuery.FieldByName("VALORGLOSA").AsFloat = 0)And(CurrentQuery.FieldByName("QTDGLOSA").AsFloat = 0)And(CurrentQuery.FieldByName("ORIGEMGLOSA").AsInteger = 1)Then
    CanContinue = False
    MsgBox("Favor informar valor ou quantidade glosada")
    Exit Sub
  End If

  If alterouRevisada Then
    If(CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S")Then
      CurrentQuery.FieldByName("USUARIOREVISOR").Value = CurrentUser
      CurrentQuery.FieldByName("DATAHORAREVISAO").Value = ServerNow
    Else
      CurrentQuery.FieldByName("USUARIOREVISOR").Clear
      CurrentQuery.FieldByName("DATAHORAREVISAO").Clear
      'se nao está revisada, nunca pode estar reconsiderada
      CurrentQuery.FieldByName("DATAHORALIBERACAO").Clear
      CurrentQuery.FieldByName("USUARIOLIBERACAO").Clear
    End If
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT ATIVA FROM SAM_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  SQL.Active = True
  If SQL.FieldByName("ATIVA").AsString = "N" Then
    MsgBox("O motivo de glosa está desativado")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  ' sms: 4142
  SQL.Clear
  SQL.Add("SELECT QTDAPRESENTADA, VALORAPRESENTADO, VALORCALCPAGTO FROM SAM_GUIA_EVENTOS WHERE HANDLE=:HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_GUIA_EVENTOS")
  SQL.Active = True

  valorcalculado = SQL.FieldByName("VALORCALCPAGTO").AsFloat
  valorapresentado = SQL.FieldByName("VALORAPRESENTADO").Value

  If SQL.FieldByName("QTDAPRESENTADA").AsFloat <CurrentQuery.FieldByName("QTDGLOSA").AsFloat Then
    MsgBox("A quantidade glosada não pode ser superior à quantidade apresentada")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If
  ' sms: 4142

  Set SQL = Nothing

  If CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat <0 Then
    CanContinue = False
    MsgBox "O valor reconsiderado deve ser maior que zero"
    Exit Sub
  End If

  If CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat <0 Then
    CanContinue = False
    MsgBox "O quantidade reconsiderada deve ser maior que zero"
    Exit Sub
  End If

  If CurrentQuery.FieldByName("GLOSADEPENDENTE").AsString = "N" Then
    'sms 50690
    If (CurrentQuery.FieldByName("ORIGEMGLOSA").AsString ="1") Then
      Dim SQLAUX As Object
      Dim vComoPagar As String
      Dim valorGlosado As Currency
      Dim vMaxGlosar As Currency
      Dim vRecebedor As Long
      Set SQLAUX = NewQuery

      SQLAUX.Clear
      SQLAUX.Add(" SELECT GE.RECEBEDOR GUIAEVENTORECEBEDOR, G.RECEBEDOR GUIARECEBEDOR, P.RECEBEDOR PEGRECEBEDOR")
      SQLAUX.Add(" FROM SAM_GUIA_EVENTOS GE ")
      SQLAUX.Add(" JOIN SAM_GUIA G ON (G.HANDLE = GE.GUIA) ")
      SQLAUX.Add(" JOIN SAM_PEG P  ON (P.HANDLE = G.PEG) ")
      SQLAUX.Add(" WHERE GE.HANDLE = :PGUIAEVENTO ")
      SQLAUX.ParamByName("PGUIAEVENTO").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
	  SQLAUX.Active = True
	  vRecebedor = IIf(Not(SQLAUX.FieldByName("GUIAEVENTORECEBEDOR").IsNull), _
                    SQLAUX.FieldByName("GUIAEVENTORECEBEDOR").AsInteger, _
                    IIf(Not(SQLAUX.FieldByName("GUIARECEBEDOR").IsNull), _
                        SQLAUX.FieldByName("GUIARECEBEDOR").AsInteger, _
                        SQLAUX.FieldByName("PEGRECEBEDOR").AsInteger))

      SQLAUX.Clear
      SQLAUX.Add(" SELECT COMOPAGARAPRESENTADONEGOCIADO FROM SAM_PRESTADOR ")
      SQLAUX.Add(" WHERE HANDLE = :PRECEBEDOR ")
      SQLAUX.ParamByName("PRECEBEDOR").AsInteger = vRecebedor
      SQLAUX.Active = True

      If SQLAUX.FieldByName("COMOPAGARAPRESENTADONEGOCIADO").AsString = "P" Then
        SQLAUX.Clear
        SQLAUX.Add("SELECT COMOPAGARAPRESENTADONEGOCIADO FROM SAM_PARAMETROSPROCCONTAS")
        SQLAUX.Active = True
      End If
      vComoPagar = SQLAUX.FieldByName("COMOPAGARAPRESENTADONEGOCIADO").AsString

      SQLAUX.Clear
      SQLAUX.Active = False
      SQLAUX.Add("SELECT SUM(VALORGLOSA) VALOR")
      SQLAUX.Add("  FROM SAM_GUIA_EVENTOS_GLOSA")
      SQLAUX.Add(" WHERE GUIAEVENTO = :GUIAEVENTO")
      SQLAUX.Add("   AND ORIGEMGLOSA = '3' ")
      SQLAUX.Add("   AND QTDGLOSA = 0 ")
      SQLAUX.ParamByName("GUIAEVENTO").AsInteger = CurrentQuery.FieldByName("GUIAEVENTO").AsInteger
      SQLAUX.Active = True
      valorGlosado = SQLAUX.FieldByName("VALOR").AsFloat
      If IsNull(valorapresentado) Then
        valorapresentado = valorcalculado
      End If
      If vComoPagar = "L" Then
        vMaxGlosar = valorapresentado - valorGlosado
      Else
        vMaxGlosar = valorcalculado '- valorGlosado
      End If
      If CurrentQuery.FieldByName("VALORGLOSA").AsCurrency > vMaxGlosar Then
        MsgBox "Valor máximo para glosar: " + Format(vMaxGlosar, "##0.00")
        CanContinue = False
        Set SQLAUX = Nothing
        Exit Sub
      End If
    End If
    Dim sqlev As Object
    Dim ValorUnitarioGlosa As Double
    Dim MaxRecon As Double
    Set sqlev = NewQuery
    sqlev.Add("SELECT VALORGLOSADO, QTDAPRESENTADA, QTDGLOSADA FROM SAM_GUIA_EVENTOS WHERE HANDLE=" + Str(RecordHandleOfTable("SAM_GUIA_EVENTOS")))
    sqlev.Active = True
    ValorUnitarioGlosa = CurrentQuery.FieldByName("VALORGLOSA").AsFloat / sqlev.FieldByName("QTDAPRESENTADA").AsFloat
    MaxRecon = ValorUnitarioGlosa * (sqlev.FieldByName("QTDAPRESENTADA").AsFloat - sqlev.FieldByName("QTDGLOSADA").AsFloat)
    If Round(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat,2) > Round(MaxRecon,2) Then ' SMS 78410 Willian - 16/3/2007
      MsgBox("Máximo para reconsiderar: " + Str(MaxRecon) + ". Para reconsiderar " + Str(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat) + ", deve-se primeiro reconsiderar As glosas dependentes")
      Set sqlev = Nothing
      CanContinue = False
      Exit Sub
    End If
    Set sqlev = Nothing
  End If
  If(arredonda(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat)>arredonda(CurrentQuery.FieldByName("QTDGLOSA").AsFloat))Or _
     (arredonda(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat)>arredonda(CurrentQuery.FieldByName("VALORGLOSA").AsFloat))Then
    MsgBox("A reconsideração não pode ser maior que a glosa")
    CanContinue = False
    Exit Sub
  End If
  If(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat >0) Or (CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat >0) Then
     If CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S" Then
       CurrentQuery.FieldByName("DATAHORALIBERACAO").Value = ServerNow
       CurrentQuery.FieldByName("USUARIOLIBERACAO").Value = CurrentUser
     Else
      CurrentQuery.FieldByName("DATAHORALIBERACAO").Clear
      CurrentQuery.FieldByName("USUARIOLIBERACAO").Clear
      CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
      CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
     End If
  Else
    'marcar o usuário revisao
    CurrentQuery.FieldByName("DATAHORALIBERACAO").Clear
    CurrentQuery.FieldByName("USUARIOLIBERACAO").Clear
    CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
    CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
  End If

  If CurrentQuery.FieldByName("GLOSAREVISADA").AsString = "S" Then
    Set SQL = NewQuery
    SQL.Add("SELECT EXIGECOMPLEMENTO FROM SAM_MOTIVOGLOSA WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active = True
    vExigeComplemento = SQL.FieldByName("EXIGECOMPLEMENTO").AsString
    Set SQL = Nothing
    If vExigeComplemento = "S" Then
      If CurrentQuery.FieldByName("COMPLEMENTO").IsNull Or _
        Trim(CurrentQuery.FieldByName("COMPLEMENTO").AsString) = "" Then
        MsgBox "Motivo de glosa exige a digitação do complemento!"
        CanContinue = False
        Exit Sub
      End If
    End If
  End If
  If CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat <>0 And CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat <>0 Then
    Set SQL = NewQuery
    If Not CurrentQuery.FieldByName("MOTIVORECONSIDERACAO").IsNull Then 'Henrique SMS: 14647
      SQL.Clear
      SQL.Active = False
      SQL.Add("SELECT EXIGECOMPLEMENTO FROM SAM_MOTIVORECONSIDERACAO WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MOTIVORECONSIDERACAO").AsInteger
      SQL.Active = True
      If SQL.FieldByName("EXIGECOMPLEMENTO").AsString = "S" Then
        If CurrentQuery.FieldByName("COMPLEMENTORECONSIDERACAO").IsNull Or _
          Trim(CurrentQuery.FieldByName("COMPLEMENTORECONSIDERACAO").AsString) = "" Then
          MsgBox "Motivo de reconsideração exige a digitação do complemento reconsideração!"
          CanContinue = False
          Exit Sub
        End If
      End If
    End If
    'If MOTIVORECONSIDERACAO.ReadOnly = False Then
    '  COMPLEMENTORECONSIDERACAO.ReadOnly = True
    '  MOTIVORECONSIDERACAO.ReadOnly = True
    '  VALORRECONSIDERADO.ReadOnly = True
    '  QTDRECONSIDERADA.ReadOnly = True
    'End If
    SQL.Active = False
    Set SQL = Nothing
  End If
  CurrentQuery.FieldByName("VALORGLOSA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("VALORGLOSA").AsFloat, "##0.00"))
  CurrentQuery.FieldByName("QTDGLOSA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("QTDGLOSA").AsFloat, "##0.00"))
  CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = CDbl(Format(CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat, "##0.00"))
  CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = CDbl(Format(CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat, "##0.00"))

End Sub

'início sms 68329 - Edilson.Castro - 19/09/2006
'alterações baseadas na sms 61368
Public Function VerificaReconsideracao As String
  Dim q1 As Object
  Dim Reconsiderar As String
  Dim vNaoReconsiderarPara As String

  Set q1=NewQuery

  VerificaReconsideracao = ""
  q1.Clear
  q1.Add("SELECT M.HANDLE ")
  q1.Add("  FROM SAM_MOTIVOGLOSA_SIS M, ")
  q1.Add("       SAM_MOTIVOGLOSA SAM, ")
  q1.Add("       SIS_MOTIVOGLOSA SIS ")
  q1.Add("WHERE SAM.HANDLE=:MOTIVOGLOSA")
  q1.Add("  AND (SIS.CODIGO IN (6010,7000,8000,8010,9650))")
  q1.Add("  AND SAM.HANDLE=M.SAMMOTIVOGLOSA")
  q1.Add("  AND SIS.HANDLE=M.SISMOTIVOGLOSA")
  q1.ParamByName("MOTIVOGLOSA").Value=CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  q1.Active=True
  If Not q1.EOF Then
    VerificaReconsideracao = "O Motivo da glosa não suporta reconsideração"
    Set q1=Nothing
    Exit Function
  End If
  Set q1=Nothing

  'Jun - SMS 44807 - 01/06/2005 - Início
  'Verificar se o motivo de glosa está inserido na carga abaixo do usuário, se não estiver então nào será possível reconsiderar a glosa

  Dim qUsuarioGlosa As Object
  Set qUsuarioGlosa = NewQuery

  qUsuarioGlosa.Clear
  qUsuarioGlosa.Add("SELECT EXIGERESPONSAVEL FROM SAM_MOTIVOGLOSA WHERE HANDLE = :HMOTIVOGLOSA")
  qUsuarioGlosa.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  qUsuarioGlosa.Active = True

  If qUsuarioGlosa.FieldByName("EXIGERESPONSAVEL").AsString = "S" Then

    Dim vQtdeUsuarioGlosa As Integer

    qUsuarioGlosa.Clear
    qUsuarioGlosa.Add("SELECT COUNT(1) QTDE")
    qUsuarioGlosa.Add("  FROM SAM_USUARIO_MOTIVOGLOSA")
    qUsuarioGlosa.Add(" WHERE USUARIO = :HUSUARIO")
    qUsuarioGlosa.Add("   AND MOTIVOGLOSA = :HMOTIVOGLOSA")
    qUsuarioGlosa.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    qUsuarioGlosa.ParamByName("HUSUARIO").AsInteger = CurrentUser
    qUsuarioGlosa.Active = True

    vQtdeUsuarioGlosa = qUsuarioGlosa.FieldByName("QTDE").AsInteger

    If vQtdeUsuarioGlosa = 0 Then
      'Agora será verificado se esta glosa está cadastrada no grupo de segurança principal do usuário corrente
      qUsuarioGlosa.Clear
      qUsuarioGlosa.Add("SELECT SUM(X.QTDE) QTDE")
      qUsuarioGlosa.Add("  FROM (")
      qUsuarioGlosa.Add("        SELECT COUNT(1) QTDE ")
      qUsuarioGlosa.Add("          FROM SAM_GRUPO_MOTIVOGLOSA   GM, ")
      qUsuarioGlosa.Add("               Z_GRUPOUSUARIOS          U")
      qUsuarioGlosa.Add("         WHERE GM.GRUPO = U.GRUPO ")
      qUsuarioGlosa.Add("           AND GM.MOTIVOGLOSA = :HMOTIVOGLOSA")
      qUsuarioGlosa.Add("           AND U.HANDLE = :HUSUARIO")
      qUsuarioGlosa.Add("         UNION ALL")
      qUsuarioGlosa.Add("        SELECT COUNT(1) QTDE ")
      qUsuarioGlosa.Add("          FROM SAM_GRUPO_MOTIVOGLOSA   GM, ")
      qUsuarioGlosa.Add("               Z_GRUPOUSUARIOGRUPOS    UG	")
      qUsuarioGlosa.Add("         WHERE UG.GRUPOADICIONADO = GM.GRUPO")
      qUsuarioGlosa.Add("           AND GM.MOTIVOGLOSA = :HMOTIVOGLOSA")
      qUsuarioGlosa.Add("           AND UG.USUARIO = :HUSUARIO")
      qUsuarioGlosa.Add("       ) X")
      qUsuarioGlosa.ParamByName("HMOTIVOGLOSA").AsInteger = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
      qUsuarioGlosa.ParamByName("HUSUARIO").AsInteger = CurrentUser
      qUsuarioGlosa.Active = True

      vQtdeUsuarioGlosa = qUsuarioGlosa.FieldByName("QTDE").AsInteger
      Set qUsuarioGlosa = Nothing

      If vQtdeUsuarioGlosa = 0 Then
        'SMS 76469 - Débora Rebello - 06/02/2007
        'Estava sendo exibida a mensagem de não permissão para o usuário mas, de acordo com o controle
        'que é feito para a habilitação ou não dos campos para a reconsideração das glosas, deve-se
        'retornar a mensagem a ser mostrada, e não mostrá-la.
        'MsgBox("Usuário não tem permissão para reconsiderar a glosa.")
         VerificaReconsideracao = "Usuário não tem permissão para reconsiderar a glosa."

        Exit Function
      End If
    Else
      Set qUsuarioGlosa = Nothing
    End If
  Else
    Set qUsuarioGlosa = Nothing
  End If

  'Jun - SMS 44807 - 01/06/2005 - Final

End Function

Public Function HabilitaCamposGlosa As Integer

    Dim SQL As Object
    Set SQL=NewQuery
    SQL.Clear
    SQL.Add("SELECT GLOSADEPENDENTE FROM SAM_MOTIVOGLOSA WHERE HANDLE=:MOTIVOGLOSA")
    SQL.ParamByName("MOTIVOGLOSA").Value=CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
    SQL.Active=True
    CurrentQuery.FieldByName("GLOSADEPENDENTE").Value=SQL.FieldByName("GLOSADEPENDENTE").AsString
    If SQL.FieldByName("GLOSADEPENDENTE").AsString = "S" Then
      CurrentQuery.FieldByName("VALORGLOSA").AsFloat = 0
      CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
      VALORGLOSA.ReadOnly = True
      'VALORRECONSIDERADO.ReadOnly = True
      CurrentQuery.FieldByName("QTDGLOSA").AsFloat = 0
	  CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
      QTDGLOSA.ReadOnly = False
      'QTDRECONSIDERADA.ReadOnly = False
    Else
      CurrentQuery.FieldByName("QTDGLOSA").AsFloat = 0
      CurrentQuery.FieldByName("QTDRECONSIDERADA").AsFloat = 0
      QTDGLOSA.ReadOnly = True
      'QTDRECONSIDERADA.ReadOnly = True
      CurrentQuery.FieldByName("VALORGLOSA").AsFloat = 0
	  CurrentQuery.FieldByName("VALORRECONSIDERADO").AsFloat = 0
      VALORGLOSA.ReadOnly = False
      'VALORRECONSIDERADO.ReadOnly = False
    End If

    Set SQL=Nothing
End Function

'fim sms 68329
