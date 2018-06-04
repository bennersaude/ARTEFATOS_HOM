'HASH: 370566108700963C55B06F97AAB688E1
'Carga 7.13.5.1.1.4.ESOCIAL_EXPORTADOS

'#Uses "*bsShowMessage"

Dim component As CSBusinessComponent
Dim CompetenciaHandle As Long
Dim CompetenciaTipoEvento As String
Dim EventoSituacao As String

Public Sub BOTAOINFORMARPROTENV_OnClick()
  If (PrepararProcesso()) Then
    component.ClearParameters
    component.AddParameter(pdtString, CompetenciaTipoEvento)
    component.AddParameter(pdtInteger, CompetenciaHandle)
    component.Execute("InformarProtocoloEnvio")
    Set component = Nothing

    RefreshCarga
  End If
End Sub

Public Function PrepararProcesso() As Boolean
  PrepararProcesso = False

  Set component = BusinessComponent.CreateInstance("Benner.Saude.eSocial.Business.BLL.EsoCompetenciaBLL, Benner.Saude.eSocial.Business")
  EventoSituacao = "4"
  CompetenciaHandle = GetHandleCompetencia()

  If (CompetenciaHandle > 0) Then
    If (ExisteEventoExportadoNaRotina(CompetenciaHandle)) Then
      CompetenciaTipoEvento = GetTipoEventoCompetencia(CompetenciaHandle)

      If (Len(CompetenciaTipoEvento) > 0) Then
        PrepararProcesso = True
      Else
        bsShowMessage("Não encontrado tipo de evento na competência.", "E")
      End If
    Else
      bsShowMessage("Não existe evento exportado na competência.", "I")
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

Public Function ExisteEventoExportadoNaRotina(HandleCompetencia As Long) As Boolean
    component.ClearParameters
    component.AddParameter(pdtInteger, HandleCompetencia)
    component.AddParameter(pdtString, EventoSituacao)
    ExisteEventoExportadoNaRotina = component.Execute("ExisteEventoComSituacaoNaCompetencia")
End Function

Public Sub RefreshCarga()
  RefreshNodesWithTable("ESO_COMPETENCIA")
End Sub
