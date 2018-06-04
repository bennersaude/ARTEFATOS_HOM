'HASH: 8564338A6AAA31642367E656ADA2E5F6
'Carga 7.13.5.1.1.2.ESOCIAL_PROCESSADOS

'#Uses "*bsShowMessage"

Dim component As CSBusinessComponent
Dim CompetenciaHandle As Long
Dim competenciaTipoEvento As String
Dim EventoSituacao As String


Public Sub BOTAOCANCELAR_OnClick()
  If (PrepararProcesso()) Then
    component.ClearParameters
    component.AddParameter(pdtString, competenciaTipoEvento)
    component.AddParameter(pdtInteger, CompetenciaHandle)
    component.AddParameter(pdtString, EventoSituacao)
    component.Execute("Cancelar")
    Set component = Nothing

    RefreshCarga
    bsShowMessage("Eventos processados da competência cancelados com sucesso.", "I")
  End If
End Sub

Public Sub BOTAOGERARXML_OnClick()
  If (PrepararProcesso()) Then
    component.ClearParameters
    component.AddParameter(pdtString, competenciaTipoEvento)
    component.AddParameter(pdtInteger, CompetenciaHandle)
    component.Execute("GerarXml")
    Set component = Nothing

    RefreshCarga
    bsShowMessage("Eventos processados da competência geraram XML com sucesso.", "I")
  End If
End Sub

Public Function PrepararProcesso() As Boolean
  PrepararProcesso = False

  Set component = BusinessComponent.CreateInstance("Benner.Saude.eSocial.Business.BLL.EsoCompetenciaBLL, Benner.Saude.eSocial.Business")
  EventoSituacao = "2"
  CompetenciaHandle = GetHandleCompetencia()

  If (CompetenciaHandle > 0) Then
    If (ExisteEventoProcessadoNaRotina(CompetenciaHandle)) Then
      competenciaTipoEvento = GetTipoEventoCompetencia(CompetenciaHandle)

      If (Len(competenciaTipoEvento) > 0) Then
        PrepararProcesso = True
      Else
        bsShowMessage("Não encontrado tipo de evento na competência.", "E")
      End If
    Else
      bsShowMessage("Não existe evento processado na competência.", "I")
    End If
  Else
    bsShowMessage("Não encontrada competência.", "E")
  End If
End Function

Public Function GetHandleCompetencia() As Long
  GetHandleCompetencia = RecordHandleOfTable("ESO_COMPETENCIA")
End Function

Public Function GetTipoEventoCompetencia(HandleCompetencia As Long) As String
  component.ClearParameters
  component.AddParameter(pdtInteger, HandleCompetencia)
  GetTipoEventoCompetencia = component.Execute("GetTipoEvento")
End Function

Public Function ExisteEventoProcessadoNaRotina(HandleCompetencia As Long) As Boolean
  component.ClearParameters
  component.AddParameter(pdtInteger, HandleCompetencia)
  component.AddParameter(pdtString, EventoSituacao)
  ExisteEventoProcessadoNaRotina = component.Execute("ExisteEventoComSituacaoNaCompetencia")
End Function

Public Sub RefreshCarga()
  RefreshNodesWithTable("ESO_COMPETENCIA")
End Sub
