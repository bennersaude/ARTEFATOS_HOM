'HASH: 2DD84C207062CEE2FA775E92854651DB
'Carga 7.13.5.1.1.3.ESOCIAL_GERADOSXML

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
    bsShowMessage("Eventos gerados XML da competência cancelados com sucesso.", "I")
  End If
End Sub

Public Sub BOTAOEXPORTARXML_OnClick()
  If (PrepararProcesso()) Then
    component.ClearParameters
    component.AddParameter(pdtString, CompetenciaTipoEvento)
    component.AddParameter(pdtInteger, CompetenciaHandle)
    component.Execute("ExportarXml")
    Set component = Nothing

    RefreshCarga
    bsShowMessage("Eventos gerados XML da competência exportados com sucesso.", "I")
  End If
End Sub

Public Function PrepararProcesso() As Boolean
  PrepararProcesso = False

  Set component = BusinessComponent.CreateInstance("Benner.Saude.eSocial.Business.BLL.EsoCompetenciaBLL, Benner.Saude.eSocial.Business")
  EventoSituacao = "3"
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
      bsShowMessage("Não existe evento para geração de XML na competência.", "I")
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

