'HASH: A716017D1197047099492F672358DCC6

Option Explicit

Public Sub Main

  On Error GoTo Exception

  Dim vsmensagem As String

  Dim RespostaStatusRecursoGlosa As Object
  Set RespostaStatusRecursoGlosa = CreateBennerObject("BENNER.SAUDE.WSTISS.GERACAOENVIOXML.ProcessamentoXml")
  vsmensagem = RespostaStatusRecursoGlosa.ProcessarRespStatusRecursoGlosa321(CurrentSystem)
  Set RespostaStatusRecursoGlosa = Nothing

  If vsmensagem <> "" Then
    InfoDescription = vsmensagem
  End If

  Exception:
    Set RespostaStatusRecursoGlosa = Nothing
    Err.Raise(Err.Number, Err, "Erro ao executar agendamento de resposta de status de recurso de glosa: " + Error)

End Sub
