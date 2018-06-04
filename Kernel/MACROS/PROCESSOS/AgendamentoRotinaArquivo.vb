'HASH: 83F7136B3A3877918215C8BFFCAEF091
'#Uses "*bsShowMessage"
'#Uses "*PrimeiroDiaCompetencia"

Public Sub Main


  Dim Interface         As Object
  Dim QAgendados        As BPesquisa
  Dim QAlteraSituacao   As BPesquisa
  Dim QCampos           As BPesquisa
  Dim QSituacao         As BPesquisa
  Dim HandleModelo      As Long
  Dim BaixaJuroMulta    As Boolean
  Dim Mensagem          As String
  Dim Retorno           As Long
  Dim motivoErro        As String
  Dim motivoErroRotina  As String
  Dim DllMessage        As String

  Set QAgendados        = NewQuery
  Set QAlteraSituacao   = NewQuery
  Set QCampos           = NewQuery
  Set QSituacao         = NewQuery

  QAgendados.Active = False
  QAgendados.Clear
  QAgendados.Add("   SELECT *                  ")
  QAgendados.Add("     FROM SFN_ROTINAARQUIVO  ")
  QAgendados.Add("    WHERE SITUACAO = 3       ")
  QAgendados.Add("    ORDER BY HANDLE          ")
  QAgendados.Active = True

  motivoErroRotina = ""

  While Not QAgendados.EOF

    motivoErro = ""
    If QAgendados.FieldByName("TABTIPO").AsInteger = 4 Then
      Dim qParam As BPesquisa
      Set qParam = NewQuery
      qParam.Clear
      qParam.Add("SELECT PERIODOFATCONINICIAL, CONTABILIZA FROM SFN_PARAMETROSFIN")
      qParam.Active = True

      If qParam.FieldByName("CONTABILIZA").AsString = "S" Then
        If (QAgendados.FieldByName("DATAINICIAL").AsDateTime >= PrimeiroDiaCompetencia(qParam.FieldByName("PERIODOFATCONINICIAL").AsDateTime)) Or ((QAgendados.FieldByName("DATAFINAL").AsDateTime >= PrimeiroDiaCompetencia(qParam.FieldByName("PERIODOFATCONINICIAL").AsDateTime))) Then
          motivoErro = "Não é permitido Processar uma rotina cuja data inicial/final esteja dento do período contábil."
          Set qParam = Nothing
          GoTo ProximaRotina
        End If
      End If

      Set qParam = Nothing
    End If

    QCampos.Active = False
    QCampos.Clear
    QCampos.Add("SELECT HANDLE ")
    QCampos.Add("  FROM SFN_MODELO_ESTRUTURA ")
    QCampos.Add(" WHERE MODELO = :MOD")
    QCampos.ParamByName("MOD").AsInteger = QAgendados.FieldByName("MODELO").AsInteger
    QCampos.Active = True
    HandleModelo = QCampos.FieldByName("HANDLE").AsInteger

    QCampos.Active = False
    QCampos.Clear
    QCampos.Add("Select COUNT(HANDLE) QT ")
    QCampos.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO ")
    QCampos.Add(" WHERE MODELOESTRUTURA = :MOD")
    QCampos.Add("   And CAMPO In (SELECT HANDLE ")
    QCampos.Add("                   FROM SIS_CONTABCAMPOS ")
    QCampos.Add("                  WHERE NOME = 'BAIXAJUROSMULTA' ")
    QCampos.Add("                     OR NOME = 'BAIXAMULTA') ")
    QCampos.ParamByName("MOD").AsInteger = HandleModelo
    QCampos.Active = True

    BaixaJuroMulta = False

    If QCampos.FieldByName("QT").AsInteger = 2 Then
      BaixaJuroMulta = True
    End If

    QCampos.Active = False
    QCampos.Clear
    QCampos.Add("Select COUNT(HANDLE) QT ")
    QCampos.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO ")
    QCampos.Add(" WHERE MODELOESTRUTURA = :MOD")
    QCampos.Add("   And CAMPO In (SELECT HANDLE ")
    QCampos.Add("                   FROM SIS_CONTABCAMPOS ")
    QCampos.Add("                  WHERE NOME = 'BAIXAJUROSMULTA' ")
    QCampos.Add("                     OR NOME = 'BAIXAJURO') ")
    QCampos.ParamByName("MOD").AsInteger = HandleModelo
    QCampos.Active = True
    If QCampos.FieldByName("QT").AsInteger = 2 Then
      BaixaJuroMulta = True
    End If

    If BaixaJuroMulta Then
      motivoErro = "O modelo não pode possuir o campo 'Baixa juros e multa' juntamente com o campo 'Baixa juro' ou com o campo 'Baixa multa'"
	  GoTo ProximaRotina
    End If

    If QAgendados.FieldByName("TABTIPO").AsInteger = 3 Then
      Dim QCheque As BPesquisa
      Set QCheque = NewQuery
      QCheque.Active = False
      QCheque.Add("SELECT NUMEROCHEQUEDISPONIVEL FROM SFN_TESOURARIA WHERE HANDLE=:TESOURARIA")
      QCheque.ParamByName("TESOURARIA").AsInteger = QAgendados.FieldByName("TESOURARIA").AsInteger
      QCheque.Active = True
      If QCheque.FieldByName("NUMEROCHEQUEDISPONIVEL").AsInteger <> QAgendados.FieldByName("NUMEROCHEQUE").AsInteger Then
        motivoErro = "Número do cheque na rotina difere da tesouraria."
        GoTo ProximaRotina
      End If
      Set QCheque = Nothing
    End If

    Set Interface = CreateBennerObject("RotArq.RotinaArquivo_ProcessaRotina")

    QAlteraSituacao.Clear
    QAlteraSituacao.Add("      UPDATE SFN_ROTINAARQUIVO SET SITUACAO = 1  ")
    QAlteraSituacao.Add("       WHERE HANDLE = :HANDLE                    ")
    QAlteraSituacao.ParamByName("HANDLE").AsInteger = QAgendados.FieldByName("HANDLE").AsInteger
    QAlteraSituacao.ExecSQL

    Retorno = Interface.Exec(CurrentSystem, _
                             QAgendados.FieldByName("HANDLE").AsInteger, _
                             0, _
                             DllMessage)

    If DllMessage <> "" Then
      motivoErro = DllMessage
      GoTo ProximaRotina
    End If

    QSituacao.Active = False
    QSituacao.Clear
    QSituacao.Add("SELECT SITUACAO ")
    QSituacao.Add("  FROM SFN_ROTINAARQUIVO ")
    QSituacao.Add(" WHERE HANDLE = :HANDLE")
    QSituacao.ParamByName("HANDLE").AsInteger = QAgendados.FieldByName("HANDLE").AsInteger
    QSituacao.Active = True

    If QSituacao.FieldByName("SITUACAO").AsInteger = 1 Then
      motivoErro = "Erro ocorrido ao processar a rotina, verificar na ocorrencia anterior a esta diretamente na rotina"
      GoTo ProximaRotina
    End If

    If QAgendados.FieldByName("SITUACAO").AsInteger = 0 Then
      ProximaRotina:

      QAlteraSituacao.Clear
      QAlteraSituacao.Add("      UPDATE SFN_ROTINAARQUIVO SET               ")
      QAlteraSituacao.Add("      SITUACAO = 3,                              ")
      QAlteraSituacao.Add("      OCORRENCIAS = :OCORRENCIAS                 ")
      QAlteraSituacao.Add("       WHERE HANDLE = :HANDLE                    ")
      QAlteraSituacao.ParamByName("HANDLE").AsInteger = QAgendados.FieldByName("HANDLE").AsInteger
      QAlteraSituacao.ParamByName("OCORRENCIAS").AsString = QAgendados.FieldByName("OCORRENCIAS").AsString + mensagemOcorrencia(motivoErro)
      QAlteraSituacao.ExecSQL
      motivoErroRotina = motivoErroRotina + "Rotina: " + QAgendados.FieldByName("HANDLE").AsString + " - Erro: " + motivoErro + Chr(13)
    End If

    QAgendados.Next
  Wend

  Set QCampos         = Nothing
  Set QSituacao       = Nothing
  Set Interface       = Nothing
  Set QAlteraSituacao = Nothing
  Set QAgendados      = Nothing

  If motivoErroRotina <> "" Then
    bsShowMessage("Operação Finalizada com erros nas seguintes rotinas arquivos:" + Chr(13) + motivoErroRotina,"E")
  End If

End Sub

Public Function mensagemOcorrencia(motivoErro As String) As String

  Dim dataAtual As Date
  Dim Mensagem  As String
  dataAtual = Now


  Mensagem = (" "                                                     + Chr(13) _
            + "=====================================================" + Chr(13) _
            + "ERRO Agendamento Rotina Arquivo"                       + Chr(13) _
	        + "=====================================================" + Chr(13) _
	        + "Usuário: Agendamento"                                  + Chr(13) _
	        + " "                                                     + Chr(13) _
            + "Início: " + Format(dataAtual, "General Date")          + Chr(13) _
	        + " "                                                     + Chr(13) _
	        + "Rotina Arquivo retornada para situação AGENDADA"       + Chr(13) _
            + "Motivo: " + motivoErro                                 + Chr(13) _
	        + " "                                                     + Chr(13) _
	        + "Fim: " + Format(dataAtual, "General Date")             + Chr(13) _
	        + " "                                                     + Chr(13) )
  mensagemOcorrencia = Mensagem


End Function
