'HASH: C50265D11D022CD0BE82B0F259325991
'Carga 7.13.5.1.1.1.ESOCIAL_INCONSISTENTES

'#Uses "*bsShowMessage"

Dim component As CSBusinessComponent
Dim CompetenciaHandle As Long
Dim CompetenciaTipoEvento As String
Dim EventoSituacao As String

Public Sub BOTAOCANCELAR_OnClick()
  If (PrepararProcesso()) Then
    component.ClearParameters
    component.AddParameter(pdtString, CompetenciaTipoEvento)
    component.AddParameter(pdtInteger, CompetenciaHandle)
    component.AddParameter(pdtString, EventoSituacao)
    component.Execute("Cancelar")
    Set component = Nothing

    RefreshCarga
    bsShowMessage("Eventos inconsistentes da competência cancelados com sucesso.", "I")
  End If
End Sub

Public Sub BOTAOREPROCESSAR_OnClick()
  If (PrepararProcesso()) Then
    component.ClearParameters
    component.AddParameter(pdtString, CompetenciaTipoEvento)
    component.AddParameter(pdtInteger, CompetenciaHandle)
    component.Execute("Reprocessar")
    Set component = Nothing

    RefreshCarga
    bsShowMessage("Eventos inconsistentes da competência reprocessados com sucesso.", "I")
  End If
End Sub

Public Function PrepararProcesso() As Boolean
  PrepararProcesso = False

  Set component = BusinessComponent.CreateInstance("Benner.Saude.eSocial.Business.BLL.EsoCompetenciaBLL, Benner.Saude.eSocial.Business")
  EventoSituacao = "1"
  CompetenciaHandle = GetHandleCompetencia()

  If (CompetenciaHandle > 0) Then
    If (ExisteEventoInconsistenteNaRotina(CompetenciaHandle)) Then
      CompetenciaTipoEvento = GetTipoEventoCompetencia(CompetenciaHandle)

      If (Len(CompetenciaTipoEvento) > 0) Then
        PrepararProcesso = True
      Else
        bsShowMessage("Não encontrado tipo de evento na competência.", "E")
      End If
    Else
      bsShowMessage("Não existe evento inconsistente na competência.", "I")
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

Public Function ExisteEventoInconsistenteNaRotina(HandleCompetencia As Long) As Boolean
  component.ClearParameters
  component.AddParameter(pdtInteger, HandleCompetencia)
  component.AddParameter(pdtString, EventoSituacao)
  ExisteEventoInconsistenteNaRotina = component.Execute("ExisteEventoComSituacaoNaCompetencia")
End Function

Public Sub RefreshCarga()
  RefreshNodesWithTable("ESO_COMPETENCIA")
End Sub
