'HASH: D77E82FA00D7949C9FFCDE4289D4B555
'#Uses "*bsShowMessage"

Public Sub Main

  Dim Interface         As Object
  Dim QAgendados        As BPesquisa
  Dim QAlteraSituacao   As BPesquisa
  Dim motivoErro        As String
  Dim motivoErroRotina  As String
  Dim DllMessage        As String
  Dim Retorno           As Long

  Set QAgendados        = NewQuery
  Set QAlteraSituacao   = NewQuery

  QAgendados.Active = False
  QAgendados.Clear
  QAgendados.Add("   SELECT *                  ")
  QAgendados.Add("     FROM SFN_ROTINADOC      ")
  QAgendados.Add("    WHERE SITUACAO = 3       ")
  QAgendados.Add("    ORDER BY HANDLE          ")
  QAgendados.Active = True

  motivoErroRotina = ""
  Set Interface = CreateBennerObject("Financeiro.RotinaDocumento_ProcessaDocumento")

  While Not QAgendados.EOF
    motivoErro = ""

    QAlteraSituacao.Clear
    QAlteraSituacao.Add("      UPDATE SFN_ROTINADOC SET                   ")
    QAlteraSituacao.Add("      SITUACAO = 1,                              ")
    QAlteraSituacao.Add("      OCORRENCIAS = :OCORRENCIAS                 ")
    QAlteraSituacao.Add("       WHERE HANDLE = :HANDLE                    ")
    QAlteraSituacao.ParamByName("HANDLE").AsInteger = QAgendados.FieldByName("HANDLE").AsInteger
    QAlteraSituacao.ParamByName("OCORRENCIAS").AsString = QAgendados.FieldByName("OCORRENCIAS").AsString + Chr(13) + Chr(13) + Chr(13) _
                                                                   + "=====================================================" + Chr(13) _
                                                                   + "Processo abaixo executado via agendamento:"
    QAlteraSituacao.ExecSQL

    Retorno = Interface.Exec(CurrentSystem, _
                             QAgendados.FieldByName("HANDLE").AsInteger, _
                             0)

    If Retorno = 1 Then
      motivoErro = "Erro no processamento da rotina Documento, verifique críticas na aba ocorrências da rotina documento"
      QAlteraSituacao.Clear
      QAlteraSituacao.Add("      UPDATE SFN_ROTINADOC     SET               ")
      QAlteraSituacao.Add("      SITUACAO = 3                               ")
      QAlteraSituacao.Add("       WHERE HANDLE = :HANDLE                    ")
      QAlteraSituacao.ParamByName("HANDLE").AsInteger = QAgendados.FieldByName("HANDLE").AsInteger
      QAlteraSituacao.ExecSQL
      motivoErroRotina = motivoErroRotina + "Rotina: " + QAgendados.FieldByName("HANDLE").AsString + " - Erro: " + motivoErro + Chr(13)

    End If

    QAgendados.Next
  Wend

  Set QAlteraSituacao = Nothing
  Set QAgendados      = Nothing
  Set Interface       = Nothing

  If motivoErroRotina <> "" Then
    bsShowMessage("Operação Finalizada com erros nas seguintes rotinas documentos:" + Chr(13) + motivoErroRotina,"E")
  End If


End Sub
