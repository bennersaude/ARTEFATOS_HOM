'HASH: 2238E42A2C5ECC6E8FA1536DC9F96CC7
Attribute VB_Name = "Module1"
'Macro da tabela SAM_PEG_RASTREADOR

Option Explicit

'#uses "*AtualizarHistoricoWorkflow"
'#uses "*FuncoesPeg"
'#uses "*GravarHistoricoRastreamentoCaixa"
'#uses "*WebServiceLogger"

Dim giHandlePegAnteEdicao   As Long
Dim giHandleCaixaAnteEdicao As Long
Dim giPrefixoRastreamento   As Integer

Public Cliente As Integer 'São os clientes atendidos pela BPO: 1 para Postal Saúde; 2 para Mapfre; 3 para FRG e 4 para Sabesprev.

Public Sub TABLE_AfterScroll

    SessionVar("HandleParaMovimentacao")    = CurrentQuery.FieldByName("HANDLE").AsString 'Variável usada na tela de movimentação.
    SessionVar("HandleParaMovimentacao_cx") = ""

    giHandlePegAnteEdicao   = CurrentQuery.FieldByName("PEG").AsInteger           'Para usar no AfterPost.
    giHandleCaixaAnteEdicao = CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger 'Para usar no AfterPost.

    If (WebVisionCode = "MOVIMENTACAO") Then
        SessionVar("K_SAM_PEG_RASTREADOR_HDL_PEG") = CStr(giHandlePegAnteEdicao)
    End If

End Sub

Public Sub TABLE_BeforeScroll

    IdentificarCliente

End Sub

Sub IdentificarCliente

    Dim qCliente  As Object
    Set qCliente  = NewQuery

    qCliente.Add("SELECT A.CODCLIENTE,                                  ")
    qCliente.Add("       B.PARAMETRO2N                                  ")
    qCliente.Add("  FROM GTO_CLIENTE A                                  ")
    qCliente.Add("  JOIN GTO_CLIENTE B ON (B.CODCLIENTE = A.PARAMETRO2N)")
    qCliente.Add(" WHERE A.PARAMETRO1S = :CLIENTEVIGENTE                ")
    qCliente.Add("   AND B.PARAMETRO1S = :RASTREAMENTO                  ")

    qCliente.ParamByName("CLIENTEVIGENTE").AsString = "CLIENTE VIGENTE"
    qCliente.ParamByName("RASTREAMENTO"  ).AsString = "RASTREAMENTO"

    qCliente.Active = True

    If (Not qCliente.EOF) Then

        Cliente               = qCliente.FieldByName("CODCLIENTE").AsInteger
        giPrefixoRastreamento = qCliente.FieldByName("PARAMETRO2N").AsInteger

    End If

    qCliente.Active = False
    Set qCliente = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

    'Esta função é chamada pelo WebService de rastreamento em
    'http://dc020-hweb1.bennercloud.com.br:81/WebAppWSTISSHom/App_Services/Benner.Saude.BRC.Rastreabilidade.Service.svc?wsdl,
    'que, após validações, grava os dados informados com esta função.

    Dim qSelect As BPesquisa
    Set qSelect = NewQuery

    '----------------------------------------------------------------
    'Para o comando K_WEBSERVICE (independentemente da visão)
    '----------------------------------------------------------------
    If (CommandID = "k_webservice") Then

         If (CurrentEntity.TransitoryVars("WS_Rastrea_Caixa").AsString <> "") And (CurrentEntity.TransitoryVars("WS_Rastrea_PEG").AsString <> "") Then

            'O rastreamento informado será atualizado no campo PEGRASTREADOR na tabela SAM_PEG.
            'O PEG informado será atualizado no campo PEG da tabela SAM_PEG_ RASTREADOR.

            On Error GoTo Finalizacao

            CurrentEntity.TransitoryVars("WS_Rastrea_Handle_lote"   ).AsInteger = 0
            CurrentEntity.TransitoryVars("WS_Rastrea_Handle_Rastrea").AsInteger = 0
            CurrentEntity.TransitoryVars("WS_Rastrea_Return"        ).AsString  = ""

            Dim viHandlePeg          As Long
            Dim viHandlePeGAtual     As Long
            Dim viHandleRastreamento As Long
            Dim viHandlePrestador    As Long
            Dim viNumeroPeg          As Long
            Dim vsCodigoRastreamento As String
            Dim vsObservacoes        As String
            Dim vsMensagemRetorno    As String

            vsMensagemRetorno = ""

            vsCodigoRastreamento = CurrentEntity.TransitoryVars("WS_Rastrea_Caixa").AsString
            viNumeroPeg          = CurrentEntity.TransitoryVars("WS_Rastrea_PEG").AsInteger

            'Valida a movimentação do histórico...
            qSelect.Clear
            qSelect.Add("SELECT HANDLE,                              ")
            qSelect.Add("       PEG,                                 ")
            qSelect.Add("       PRESTADOR,                           ")
            qSelect.Add("       OBSERVACAO                           ")
            qSelect.Add("  FROM SAM_PEG_RASTREADOR                   ")
            qSelect.Add(" WHERE CODRASTREAMENTO = :CODIGORASTREAMENTO")

            qSelect.ParamByName("CODIGORASTREAMENTO").AsString = vsCodigoRastreamento

            qSelect.Active = True

            If (qSelect.EOF) Then

               vsMensagemRetorno = "Não foi encontrado um rastreamento de PEG com o código " + vsCodigoRastreamento + "."
               GoTo Finalizacao

            Else

               vsObservacoes = Trim(qSelect.FieldByName("OBSERVACAO").AsString) + " - PEG " + CStr(viNumeroPeg) + " em " + Format(ServerNow, "dd/mm/yyyy hh:nn:ss ") + "."

               If (Len(Trim(vsObservacoes)) >= 4000) Then

                   vsMensagemRetorno = "O texto das observações ultrapassa os 4.000 caracteres suportados."
                   GoTo Finalizacao

               End If

            End If

            CurrentEntity.TransitoryVars("WS_Rastrea_Handle_Rastrea").AsInteger = qSelect.FieldByName("HANDLE").AsInteger

            viHandlePeGAtual  = qSelect.FieldByName("PEG").AsInteger
            viHandlePrestador = qSelect.FieldByName("PRESTADOR").AsInteger

            If (CurrentEntity.TransitoryVars("WS_Rastrea_Handle_Rastrea").AsInteger = 0) Then

                vsMensagemRetorno = "O identificador do rastreamento é inválido."
                GoTo Finalizacao

            End If

            qSelect.Active = False

            qSelect.Clear
            qSelect.Add("SELECT MAX(HANDLE) ULTIMOINCLUIDO      ")
            qSelect.Add("  FROM WFL_HISTORICO_LOTE              ")
            qSelect.Add(" WHERE RASTREADOR = :HANDLERASTREAMENTO")
            qSelect.Add("   AND DATASAIDA IS NULL               ")

            qSelect.ParamByName("HANDLERASTREAMENTO").AsInteger = CurrentEntity.TransitoryVars("WS_Rastrea_Handle_Rastrea").AsInteger

            qSelect.Active = True

            If (Not qSelect.EOF) Then
               CurrentEntity.TransitoryVars("WS_Rastrea_Handle_lote").AsInteger = qSelect.FieldByName("ULTIMOINCLUIDO").AsInteger
            End If

            qSelect.Active = False

            qSelect.Clear
            qSelect.Add("SELECT HANDLE,                     ")
            qSelect.Add("       PEGRASTREADOR               ")
            qSelect.Add("  FROM SAM_PEG                     ")
            qSelect.Add(" WHERE PEG       = :NUMEROPEG      ")
            qSelect.Add("   AND RECEBEDOR = :HANDLERECEBEDOR")

            qSelect.ParamByName("NUMEROPEG"      ).AsInteger = viNumeroPeg
            qSelect.ParamByName("HANDLERECEBEDOR").AsInteger = viHandlePrestador

            qSelect.Active = True

            If (qSelect.EOF) Then

                qSelect.Active = False

                qSelect.Clear
                qSelect.Add("SELECT PRESTADOR,      ")
                qSelect.Add("       NOME            ")
                qSelect.Add("  FROM SAM_PRESTADOR   ")
                qSelect.Add(" WHERE HANDLE = :HANDLE")

                qSelect.ParamByName("HANDLE").AsInteger = viHandlePrestador

                  qSelect.Active = True

                vsMensagemRetorno = "O PEG " + CStr(viNumeroPeg) + " não foi encontrado para o prestador " + _
                                    qSelect.FieldByName("PRESTADOR").AsString + " - " + qSelect.FieldByName("PRESTADOR").AsString + "."
                GoTo Finalizacao

            End If

            viHandlePeg          = qSelect.FieldByName("HANDLE").AsInteger
            viHandleRastreamento = qSelect.FieldByName("PEGRASTREADOR").AsInteger

            vsMensagemRetorno = AtualizarHistoricoWorkflow(CurrentEntity.TransitoryVars("WS_Rastrea_Handle_lote").AsInteger, _
                                                           CurrentEntity.TransitoryVars("WS_Rastrea_Handle_Rastrea").AsInteger, _
                                                           "BRCTRI05", _
                                                           1, _
                                                           1, _
                                                           "")
            Dim qUpdate As Object
            Set qUpdate = NewQuery

              If ((UCase(vsMensagemRetorno) = "OK") Or (vsMensagemRetorno = "")) Then

                  AtualizarRastreamento _
                      vsCodigoRastreamento, _
                      vsObservacoes, _
                      viHandlePeg

            End If

              If (UCase(vsMensagemRetorno) = UCase("A regra não permite transição para o status solicitado.")) Then

                If ((viHandlePeGAtual = 0) Or (viHandlePeGAtual <> viHandlePeg)) And (viHandleRastreamento = 0) Then

                    AtualizarRastreamento _
                          vsCodigoRastreamento, _
                          vsObservacoes, _
                          viHandlePeg

                   vsMensagemRetorno = "OK"

               End If

            End If

Finalizacao:

            If (Err.Description <> "") Then

                If (vsMensagemRetorno <> "") Then
                    vsMensagemRetorno = vsMensagemRetorno + Chr(13) + Chr(10)
                End If

                vsMensagemRetorno = vsMensagemRetorno + "*Erro: " + Err.Description

            End If

            CurrentEntity.TransitoryVars("WS_Rastrea_Return").AsString = vsMensagemRetorno
            CurrentEntity.TransitoryVars("WS_Rastrea_DEBUG").AsString  = ""

         End If

    End If

    '----------------------------------------------------------------
    'Para a visão TRIAGEM2
    '----------------------------------------------------------------
    If (WebVisionCode = "TRIAGEM2") Then

        If (CommandID = "K_IMPR_TRIAGEM") Then

            qSelect.Active = False

            qSelect.Clear
            qSelect.Add("SELECT HANDLE                          ")
            qSelect.Add("  FROM R_RELATORIOS                    ")
            qSelect.Add(" WHERE UPPER(CODIGO) = :CODIGORELATORIO")

            qSelect.ParamByName("CODIGORELATORIO").AsString = "K-010"

            qSelect.Active = True

            If (qSelect.EOF) Then

                CancelDescription = "Não foi encontrado o relatório com o código K-010."
                CanContinue = False

                qSelect.Active = False
                Set qSelect = Nothing

                Exit Sub

            Else

                Dim voRelatorio As CSReportPrinter
                Set voRelatorio = NewReport(qSelect.FieldByName("HANDLE").AsInteger)

                SessionVar("ProtocoloTriagem") = CurrentQuery.FieldByName("HANDLE").AsString

                voRelatorio.Preview

                Set voRelatorio = Nothing

                InfoDescription = "O protocolo de triagem foi impresso."

            End If

        End If

    End If

    '----------------------------------------------------------------
    'Para a visão REMESSA
    '----------------------------------------------------------------
    If (WebVisionCode = "REMESSA") Then

        '------------------------------------------------------------
        'Para o comando K_IMPR_REMESSA
        '------------------------------------------------------------
        If (CommandID = "K_IMPR_REMESSA") Then

            Dim qr As Object
            Set qr = NewQuery

            qr.Add("SELECT HANDLE FROM R_RELATORIOS WHERE UPPER(CODIGO) = 'K-008' ")
            qr.Active = True

            If Not qr.EOF Then

                  Dim rep2 As CSReportPrinter

                  Set rep2 = NewReport(qr.FieldByName("HANDLE").AsInteger)


                  SessionVar("ProtocoloRemessa") = CurrentQuery.FieldByName("HANDLE").AsString

                  rep2.Preview
                  Set rep2 = Nothing

                  InfoDescription = "Protocolo de Remessa impresso com sucesso!"

            Else
                  CanContinue = False
                  CancelDescription = "Não encontrado relatório com código K-008."

            End If

            Set qr = Nothing

        End If

        '------------------------------------------------------------
        'Para o comando K_ENVIAR_REMESSA
        '------------------------------------------------------------
        If (CommandID = "K_ENVIAR_REMESSA") Then

            vsMensagemRetorno = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, _
                                                           CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                           "POSREM02", _
                                                           1, _
                                                           1, _
                                                           "")
            InfoDescription = "A remessa foi enviada."

        End If

        '------------------------------------------------------------
        'Para o comando K_CANC_REM
        '------------------------------------------------------------
        If (CommandID = "K_CANC_REM") Then

            vsMensagemRetorno = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, _
                                                           CurrentQuery.FieldByName("HANDLE").AsInteger, "POSREM03", _
                                                           1, _
                                                           1, _
                                                           "")
            InfoDescription = "A remessa foi cancelada."

        End If

    End If

    qSelect.Active = Nothing
    Set qSelect = Nothing

End Sub

Sub AtualizarRastreamento(psCodigoRastreamento As String, psObservacoes As String, piHandlePeg As Long)

    Dim qSelect As BPesquisa
    Set qSelect = NewQuery

    qSelect.Add("SELECT PEG                                  ")
    qSelect.Add("  FROM SAM_PEG_RASTREADOR                   ")
    qSelect.Add(" WHERE CODRASTREAMENTO = :CODIGORASTREAMENTO")

    qSelect.ParamByName("CODRASTREAMENTO").AsString = psCodigoRastreamento

    qSelect.Active = True

    If (Not InTransaction) Then StartTransaction

    AtualizarRastreamentoPeg _
        psCodigo         := psCodigoRastreamento, _
        piHandlePeg      := piHandlePeg, _
        piHandlePegAtual := qSelect.FieldByName("PEG").AsInteger, _
        psObservacoes    := psObservacoes

    AtualizarPeg(piHandlePeg)

    If (InTransaction) Then Commit

    qSelect.Active = False

    Set qSelect = Nothing

    AtualizarDataPagamento _
        piHandlePeg                  := piHandlePeg, _
        psOrigem                     := "VerificaRastreabilidade", _
        pbAtualizarTambemRecebimento := True

End Sub

Sub AtualizarRastreamentoPeg(psCodigo As String, piHandlePeg As Long, piHandlePegAtual As Long, psObservacoes As String)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    qUpdate.Add("UPDATE SAM_PEG_RASTREADOR       ")
    qUpdate.Add("   SET OBSERVACAO = :OBSERVACOES")

    If (piHandlePegAtual = 0) Then

        qUpdate.Add("       ,PEG       = :PEG")
        qUpdate.ParamByName("PEG").AsInteger = piHandlePeg

    End If

    qUpdate.Add(" WHERE CODRASTREAMENTO = :CODIGORASTREAMENTO")

    qUpdate.ParamByName("OBSERVACOES"       ).AsString = psObservacoes
    qUpdate.ParamByName("CODIGORASTREAMENTO").AsString = psCodigo

    qUpdate.ExecSQL

    Set qUpdate = Nothing

End Sub

Sub AtualizarPeg(piHandlePeg As Long)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    qUpdate.Add("UPDATE SAM_PEG                            ")
    qUpdate.Add("   SET PEGRASTREADOR = :HANDLERASTREAMENTO")
    qUpdate.Add(" WHERE HANDLE = :HANDLE                   ")

    qUpdate.ParamByName("HANDLE"            ).AsInteger = piHandlePeg
    qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = CurrentEntity.TransitoryVars("WS_Rastrea_Handle_Rastrea").AsInteger

    qUpdate.ExecSQL

    Set qUpdate = Nothing

End Sub

Public Sub TABLE_NewRecord

    If ((WebVisionCode = "REMESSA") Or (WebVisionCode = "TRIAGEM2")) Then

        CurrentQuery.FieldByName("HANDLE"         ).AsInteger = NewHandle("K_SAM_PEG_RASTREADOR")
        CurrentQuery.FieldByName("CODRASTREAMENTO").AsString  = Trim(Format(CurrentQuery.FieldByName("HANDLE").AsString, CStr(giPrefixoRastreamento) + "00000000"))

    End If

    If (WebVisionCode = "TRIAGEM2") Then

        If (IsNumeric(SessionVar("ultima_caixa_digitada"))) Then

            CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger = CLng(SessionVar("ultima_caixa_digitada"))

            If (CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger <= 0) Then
                CurrentQuery.FieldByName("ARQUIVOFISICO").AsString = ""
            End If

        End If

    End If

End Sub

Function ObterOrigemPeloCodigo(piCodigo As Long)

    Dim qOrigem As BPesquisa
    Set qOrigem = NewQuery

    qOrigem.Add("SELECT HANDLE                ")
    qOrigem.Add("  FROM SAM_ORIGEMRASTREAMENTO")
    qOrigem.Add(" WHERE CODORIGEM = :CODIGO   ")

    qOrigem.ParamByName("CODIGO").AsInteger = piCodigo
    qOrigem.Active = True

    If (qOrigem.EOF) Then
        ObterOrigemPeloCodigo = 0
    Else
        ObterOrigemPeloCodigo = qOrigem.ParamByName("HANDLE").AsInteger
    End If

    qOrigem.Active = False
    Set qOrigem = Nothing

End Function


Public Sub TABLE_BeforePost(CanContinue As Boolean)

    Dim vsMensagem                 As String
    Dim viHandleOrigemRegional     As Long
    Dim viHandleOrigemPrestador    As Long
    Dim viHandleOrigemBeneficiario As Long

    vsMensagem                 = ""
    viHandleOrigemRegional     = ObterOrigemPeloCodigo(3)
    viHandleOrigemPrestador    = ObterOrigemPeloCodigo(5)
    viHandleOrigemBeneficiario = ObterOrigemPeloCodigo(9)

    '----------------------------------------------------------------
    'Para as visões REMESSA e TRIAGEM2
    '----------------------------------------------------------------
    If ((WebVisionCode = "REMESSA") Or (WebVisionCode = "TRIAGEM2")) Then

        If (viHandleOrigemRegional = 0) Then

            CancelDescription = "Não existe uma origem parametrizada para regional (código 3) no sistema."
            CanContinue = False
            Exit Sub

        End If

        If (CurrentQuery.FieldByName("QTDGUIAS").AsInteger <= 0) Then

            CancelDescription = "A quantidade de guias deve ser um valor maior que zero."
            CanContinue = False
            Exit Sub

        End If

    End If

    '----------------------------------------------------------------
    'Para as visões RETORNO e TRIAGEM2
    '----------------------------------------------------------------
    If ((WebVisionCode = "RETORNO") Or (WebVisionCode = "TRIAGEM2")) Then

        If (((CurrentQuery.FieldByName("PEG").IsNull) And (Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull)) And _
            ((WebVisionCode = "RETORNO") Or _
             ((CurrentQuery.FieldByName("ORIGEM").AsInteger <> viHandleOrigemPrestador) And (Not CurrentQuery.FieldByName("ENVIAVERIFICACAO").AsBoolean)))) Then

            CancelDescription = "Selecione um PEG para a caixa informada."
            CanContinue = False
            Exit Sub

        End If

    End If

    '----------------------------------------------------------------
    'Para a visão REMESSA
    '----------------------------------------------------------------
    If (WebVisionCode = "REMESSA") Then

        If (viHandleOrigemRegional = 0) Then

            CancelDescription = "Não existe uma origem parametrizada como 'Regional' (código 3) no sistema."
            CanContinue = False
            Exit Sub

        End If

        CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemRegional

        If ((CurrentQuery.FieldByName("TIPOREMESSA").AsInteger = 1) And (CurrentQuery.FieldByName("PRESTADOR").IsNull)) Then

            CancelDescription = "Prestador é obrigatório para remessa de faturamento."
            CanContinue = False
            Exit Sub

        End If

        If (CurrentQuery.FieldByName("TIPOREMESSA").AsInteger = 2) Then

            If (CurrentQuery.FieldByName("BENEFICIARIO").IsNull) Then

                CancelDescription = "Beneficiário é obrigatório para remessa de reembolso."
                CanContinue = False
                Exit Sub

            End If

            If (Not CurrentQuery.FieldByName("PRESTADOR").IsNull) Then

                CancelDescription = "Não se pode informar o prestador em remessas de reembolso."
                CanContinue = False
                Exit Sub

            End If

        Else

            If (Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull) Then

                CancelDescription = "Não se pode informar o beneficiário em remessas que não sejam de reembolsoO."
                CanContinue = False
                Exit Sub

            End If

        End If

        If (CurrentQuery.FieldByName("DATARECEPCAO").AsDateTime > ServerNow) Then

            CancelDescription = "A data de recepção não pode ser posterior à data atual."
            CanContinue = False
            Exit Sub

        End If

        If ((CurrentQuery.FieldByName("VLRAPRESENTADO").IsNull) And (Not CurrentQuery.FieldByName("VLRNOTAFISCAL").IsNull)) Then
            CurrentQuery.FieldByName("VLRAPRESENTADO").AsFloat = CurrentQuery.FieldByName("VLRNOTAFISCAL").AsFloat
        End If

        If (CurrentQuery.FieldByName("TIPOREMESSA").AsInteger <> 4) Then 'O tipo de remessa é "Não se Aplica".

            vsMensagem = ValidarRegra

            If (vsMensagem <> "") Then

                CancelDescription = vsMensagem
                CanContinue = False
                Exit Sub

            End If

        End If

    End If

    '----------------------------------------------------------------
    'Para a visão TRIAGEM2
    '----------------------------------------------------------------
    If (WebVisionCode = "TRIAGEM2") Then

        If ((CurrentQuery.State = 3) And (Not CurrentQuery.FieldByName("HISTORICOWORKFLOW").IsNull)) Then 'Registro em inclusão sem histórico do workflow informado.

            CancelDescription = "Este PEG já está rastreado."
            CanContinue = False
            Exit Sub

        End If

        If (CurrentQuery.FieldByName("ORIGEM").IsNull) Then

            CancelDescription = "Informe a origem."
            CanContinue = False
            Exit Sub

        End If

        Dim qSelect As BPesquisa
        Set qSelect = NewQuery

        If (Not CurrentQuery.FieldByName("PEG").IsNull) Then

            qSelect.Active = False

            qSelect.Clear
            qSelect.Add("SELECT CODRASTREAMENTO          ")
            qSelect.Add("  FROM SAM_PEG_RASTREADOR       ")
            qSelect.Add(" WHERE PEG     = :HANDLEPEG     ")
            qSelect.Add("   AND HANDLE <> :HANDLECORRENTE")

            qSelect.ParamByName("HANDLEPEG"     ).AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
            qSelect.ParamByName("HANDLECORRENTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

            qSelect.Active = True

            If (Not qSelect.EOF) Then

                CancelDescription = "Este PEG já está vinculado ao rastreamento " + qSelect.FieldByName("CODRASTREAMENTO").AsString + "."
                CanContinue = False

                qSelect.Active = False
                Set qSelect = Nothing

                Exit Sub

            End If

        End If

        IdentificarCliente

        If ((Not CurrentQuery.FieldByName("PEG").IsNull) And (Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull)) Then

            qSelect.Active = False

            qSelect.Clear
            qSelect.Add("SELECT P.PEGARQUIVOFISICO,                                      ")
            qSelect.Add("       F.DESCRICAO        CAIXA                                 ")
            qSelect.Add("  FROM SAM_PEG              P                                   ")
            qSelect.Add("  JOIN SAM_PEGARQUIVOFISICO F ON (F.HANDLE = P.PEGARQUIVOFISICO)")
            qSelect.Add(" WHERE P.HANDLE  = :HANDLEPEG                                   ")
            qSelect.Add("   AND F.HANDLE <> :ARQUIVOFISICO                               ")

            qSelect.ParamByName("HANDLEPEG"    ).AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
            qSelect.ParamByName("ARQUIVOFISICO").AsInteger = CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger

            qSelect.Active = True

            If ((Not qSelect.EOF) And (qSelect.FieldByName("PEGARQUIVOFISICO").AsFloat > 0)) Then

                CancelDescription = "Este PEG já pertence à caixa nº " + qSelect.FieldByName("CAIXA").AsString + "."
                CanContinue = False

                qSelect.Active = False
                Set qSelect = Nothing

                Exit Sub

              ElseIf (Not OrigemDemandaInclusaoEmCaixa(CurrentQuery.FieldByName("ORIGEM").AsInteger)) Then

                qSelect.Active = False

                qSelect.Clear
                qSelect.Add("SELECT NOMEORIGEM                    ")
                qSelect.Add("  FROM SAM_ORIGEMRASTREAMENTO        ")
                qSelect.Add(" WHERE STATUS = :SIM                 ")
                qSelect.Add("   AND CODORIGEM  IN (2, 6, 7, 9, 10)")

                qSelect.ParamByName("SIM").AsString = "S"

                qSelect.Active = True

                Dim vsOrigensPermitidas   As String
                Dim vsDescricaoUltimoItem As String

                vsOrigensPermitidas = ""

                While (Not qSelect.EOF)

                    vsOrigensPermitidas = vsOrigensPermitidas + "'" + qSelect.FieldByName("NOMEORIGEM").AsString + "'"

                    qSelect.Next 'Olha o próximo.

                    If (Not qSelect.EOF) Then 'Ainda tem mais coisa...

                        vsDescricaoUltimoItem = qSelect.FieldByName("NOMEORIGEM").AsString

                        qSelect.Next 'Então, olha mais um à frente pra ver se coloca um "ou" ou uma vírgula.

                        If (qSelect.EOF) Then
                            vsOrigensPermitidas = vsOrigensPermitidas + " ou '" + vsDescricaoUltimoItem + "'"
                        Else
                            vsOrigensPermitidas = vsOrigensPermitidas + ", "
                            qSelect.Prior 'Volta pro anterior, pois só foi ali na frente pra ver se era o último.
                        End If

                    End If

                Wend

                If (vsOrigensPermitidas = "") Then
                    CancelDescription = "Não existem origens para rastreamento que estejam ativas e que demandem inclusão de PEGs em caixas."
                Else
                    CancelDescription = "Incluir na caixa somente PEGs cuja origem seja " + vsOrigensPermitidas + "."
                End If

                CanContinue = False

                  qSelect.Active = False
                Set qSelect = Nothing

                Exit Sub

            End If

        End If

        qSelect.Active = False
        Set qSelect = Nothing

        If (viHandleOrigemPrestador = 0) Then

            CancelDescription = "Não existe uma origem parametrizada para prestador (código 5) no sistema."
            CanContinue = False
            Exit Sub

        End If

        If (viHandleOrigemBeneficiario = 0) Then

            CancelDescription = "Não existe uma origem parametrizada para beneficiário (código 9) no sistema."
            CanContinue = False
            Exit Sub

        End If

        If ((CurrentQuery.FieldByName("PEG").IsNull) And _
            (CurrentQuery.FieldByName("PRESTADOR").IsNull) And _
            (CurrentQuery.FieldByName("ORIGEM").AsInteger <> viHandleOrigemBeneficiario)) Then

            CancelDescription = "Se não há PEG informado, é necessário informar o prestador."
              CanContinue = False
            Exit Sub

        End If

        If ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemBeneficiario) And (CurrentQuery.FieldByName("BENEFICIARIO").IsNull)) Then

            CancelDescription = "É necessário informar o beneficiário."
              CanContinue = False
            Exit Sub

        End If

        If (CurrentQuery.FieldByName("DATARECEPCAO").AsDateTime > ServerNow) Then

            CancelDescription = "A data da recepção não pode ser futura."
              CanContinue = False
            Exit Sub

        End If

        If (CurrentQuery.FieldByName("VLRAPRESENTADO").IsNull) And (Not CurrentQuery.FieldByName("VLRNOTAFISCAL").IsNull) Then
            CurrentQuery.FieldByName("VLRAPRESENTADO").AsFloat = CurrentQuery.FieldByName("VLRNOTAFISCAL").AsFloat
        End If

        If  ((CurrentQuery.FieldByName("ORIGEM").AsInteger <> viHandleOrigemPrestador) And _
             (CurrentQuery.FieldByName("ORIGEM").AsInteger <> viHandleOrigemRegional) And _
             (Not CurrentQuery.FieldByName("ENVIAVERIFICACAO").AsBoolean)) Then

            vsMensagem = ValidarRegra

            If (vsMensagem <> "") Then

                CancelDescription = vsMensagem
                CanContinue = False
                Exit Sub

            End If

        End If

    End If

    '----------------------------------------------------------------
    'Para a visão RETORNO
    '----------------------------------------------------------------
    If (WebVisionCode = "RETORNO")  Then

        If (CurrentQuery.FieldByName("ENVIAVERIFICACAO").AsBoolean = Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull) Then

            CancelDescription = "Selecione uma caixa ou faça enviar para verificação"

            If (CurrentQuery.FieldByName("ENVIAVERIFICACAO").AsBoolean) Then
                CancelDescription = CancelDescription + ", mas não é permitido fazer ambos"
            End If

            CancelDescription = CancelDescription + "."

            CanContinue = False
            Exit Sub

        End If

    End If

End Sub

Function OrigemDemandaInclusaoEmCaixa(piHandleOrigem As Long) As Boolean

    'Deve incluir em caixas somente:
    '    2  - Orizon,
    '    6  - TRIX,
    '    9  - Beneficiário,
    '    7  - Prestador Eletrônico
    '    10 - Odontologia

    Dim qOrigem As BPesquisa
    Set qOrigem = NewQuery

    qOrigem.Add("SELECT 1                            ")
    qOrigem.Add("  FROM SAM_ORIGEMRASTREAMENTO       ")
    qOrigem.Add(" WHERE CODORIGEM IN (2, 6, 7, 9, 10)")
    qOrigem.Add("   AND HANDLE = :HANDLE             ")

    qOrigem.ParamByName("HANDLE").AsInteger = piHandleOrigem

    qOrigem.Active = True

    OrigemDemandaInclusaoEmCaixa = Not qOrigem.EOF

    qOrigem.Active = False
    Set qOrigem = Nothing

End Function

Public Function ValidarRegra As String

    Dim viHandle As Long
    Dim qSelect  As Object

    Set qSelect = NewQuery

    ValidarRegra = ""

    '----------------------------------------------------------------
    'Para a visão REMESSA
    '----------------------------------------------------------------
    If (WebVisionCode = "REMESSA") Then

        qSelect.Active = False

        qSelect.Clear

        If ((CurrentQuery.FieldByName("TIPOREMESSA").AsInteger = 1) Or (CurrentQuery.FieldByName("TIPOREMESSA").AsInteger = 3)) Then 'Remessa de faturamento ou de recurso de glosa.

            qSelect.Add("SELECT FILIALPADRAO HANDLEFILIAL")
            qSelect.Add("  FROM SAM_PRESTADOR            ")

            viHandle = CurrentQuery.FieldByName("PRESTADOR").AsInteger

        End If

        If (CurrentQuery.FieldByName("TIPOREMESSA").AsFloat = 2)  Then 'Remessa de reembolso.

            qSelect.Add("SELECT FILIALCUSTO HANDLEFILIAL")
            qSelect.Add("  FROM SAM_BENEFICIARIO        ")

            viHandle = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger

        End If

        qSelect.Add(" WHERE HANDLE = :HANDLE")
        qSelect.ParamByName("HANDLE").AsInteger = viHandle
        qSelect.Active = True

        If (Not qSelect.EOF) Then
            CurrentQuery.FieldByName("FILIALPADRAO").AsInteger = qSelect.FieldByName("HANDLEFILIAL").AsInteger
        End If

        qSelect.Active = False

    End If

    '----------------------------------------------------------------
    'Para a visão TRIAGEM2
    '----------------------------------------------------------------
    If (WebVisionCode = "TRIAGEM2") Then

        Dim viHandleOrigemPrestadorManual As Long
          Dim viHandleOrigemBeneficiario    As Long

        viHandleOrigemPrestadorManual = ObterOrigemPeloCodigo(8)
        viHandleOrigemBeneficiario    = ObterOrigemPeloCodigo(9)

        If  ((CurrentQuery.FieldByName("ORIGEM").AsInteger <> viHandleOrigemPrestadorManual) And _
             (CurrentQuery.FieldByName("ORIGEM").AsInteger <> viHandleOrigemBeneficiario)) Then

            qSelect.Active = False

            qSelect.Clear
            qSelect.Add("SELECT P.RECEBEDOR,                               ")
            qSelect.Add("       P.IDENTIFICADORLOTE,                       ")
            qSelect.Add("       R.FILIALPADRAO                             ")
            qSelect.Add("  FROM SAM_PEG       P                            ")
            qSelect.Add("  JOIN SAM_PRESTADOR R ON (R.HANDLE = P.RECEBEDOR)")
            qSelect.Add(" WHERE P.HANDLE = :HANDLEPEG                      ")

            qSelect.ParamByName("HANDLEPEG").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger

            qSelect.Active = True

            If (Not qSelect.EOF) Then

                CurrentQuery.FieldByName("PRESTADOR"    ).AsInteger = qSelect.FieldByName("RECEBEDOR").AsInteger
                CurrentQuery.FieldByName("LOTEPRESTADOR").AsString  = Trim(qSelect.FieldByName("IDENTIFICADORLOTE").AsString)
                CurrentQuery.FieldByName("FILIALPADRAO" ).AsInteger = qSelect.FieldByName("FILIALPADRAO").AsInteger

                If (qSelect.FieldByName("FILIALPADRAO").IsNull) Then
                    ValidarRegra = "O recebedor do PEG não tem uma filial padrão definida."
                End If

            Else
                ValidarRegra = "PEG ou lote não encontrado."
            End If

        End If

    End If

    qSelect.Active = False
    Set qSelect = Nothing

End Function

Public Sub TABLE_AfterPost

    Dim vsMensagemRetornada As String

    '----------------------------------------------------------------
    'Para as visões RECEPCAO e REMESSA
    '----------------------------------------------------------------
    If ((WebVisionCode = "RECEPCAO") Or (WebVisionCode = "REMESSA")) Then

        vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "POSREM01", 1, 1, "") 'PROTOCOLO GERADO NA ABERTURA

        If (vsMensagemRetornada = "OK") Then
            InfoDescription = "O registro foi incluído."
        Else
            InfoDescription = vsMensagemRetornada
        End If

    End If

    '----------------------------------------------------------------
    'Para a visão RETORNO
    '----------------------------------------------------------------
    If (WebVisionCode = "RETORNO")  Then

        If (CurrentQuery.FieldByName("ENVIAVERIFICACAO").AsBoolean) Then
            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI07", 1, 1, "") 'DIGITAÇÃO VERIFICAR
        End If

        If (Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull) Then
            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI06", 1, 1, "") 'DIGITAÇÃO OK
        End If

        If (vsMensagemRetornada = "OK") Then

            If (Not CurrentQuery.FieldByName("PEG").IsNull) Then

                Atualizar _
                    piHandleRastreamento := CurrentQuery.FieldByName("HANDLE").AsInteger, _
                    piHandlePeg          := giHandlePegAnteEdicao, _
                    piHandleCaixa        := giHandleCaixaAnteEdicao

            End If

            InfoDescription = InfoDescription + Chr(13) + Chr(10) + "O registro foi incluído."

        Else
            InfoDescription = InfoDescription + Chr(13) + Chr(10) + vsMensagemRetornada
        End If

    End If

    '----------------------------------------------------------------
    'Para a visão TRIAGEM2
    '----------------------------------------------------------------
    If (WebVisionCode = "TRIAGEM2") Then

        Dim viHandleOrigemRegional        As Long
        Dim viHandleOrigemPrestador       As Long
        Dim viHandleOrigemPrestadorManual As Long
          Dim viHandleOrigemBeneficiario    As Long

        viHandleOrigemRegional        = ObterOrigemPeloCodigo(3)
        viHandleOrigemPrestador       = ObterOrigemPeloCodigo(5)
        viHandleOrigemPrestadorManual = ObterOrigemPeloCodigo(8)
        viHandleOrigemBeneficiario    = ObterOrigemPeloCodigo(9)

        If (CurrentQuery.FieldByName("ENVIAVERIFICACAO").AsBoolean) Then

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI02", 1, 1, "")

        ElseIf ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemPrestador) And (CurrentQuery.FieldByName("PEG").IsNull)) Then

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI04", 1, 1, "") 'ENV. DIG. EXTERNA

        ElseIf ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemRegional) And (CurrentQuery.FieldByName("PEG").IsNull)) Then

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI04", 1, 1, "") 'ENV. DIG. EXTERNA

        ElseIf ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemPrestadorManual) And (CurrentQuery.FieldByName("PEG").IsNull) And (Cliente <> 4)) Then 'Quando o cliente não for a Sabesprev.

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI04", 1, 1, "") 'ENV. DIG. EXTERNA

        ElseIf ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemPrestadorManual) And (CurrentQuery.FieldByName("PEG").IsNull) And (Cliente = 4)) Then 'Somente quando o cliente for a Sabesprev.

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCPRC07", 1, 1, "") 'ENV. DIG. EXTERNA

        ElseIf ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemBeneficiario) And (CurrentQuery.FieldByName("PEG").IsNull)) Then

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCPRC07", 1, 1, "") 'ENV. DIG. EXTERNA

        ElseIf ((CurrentQuery.FieldByName("ORIGEM").AsInteger = viHandleOrigemBeneficiario) And (Not CurrentQuery.FieldByName("PEG").IsNull)) Then

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCPRC01", 1, 1, "")

        ElseIf (CurrentQuery.FieldByName("PEG").IsNull) Then

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI02", 1, 1, "") 'VERIFICAR PENDÊNCIA

        Else

            vsMensagemRetornada = AtualizarHistoricoWorkflow(CurrentQuery.FieldByName("HISTORICOWORKFLOW").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, "BRCTRI01", 1, 1, "") 'TRIAGEM OK

        End If

        Dim vbAtualizouDataPagamento

        vbAtualizouDataPagamento = False

        If (vsMensagemRetornada = "OK") Then

               Atualizar _
                piHandleRastreamento := CurrentQuery.FieldByName("HANDLE").AsInteger, _
                piHandlePeg          := giHandlePegAnteEdicao, _
                piHandleCaixa        := giHandleCaixaAnteEdicao

            If (Cliente = 1) Then 'Somente quando o cliente for a Postal Saúde.

                If (Not CurrentQuery.FieldByName("PEG").IsNull) Then

                    vbAtualizouDataPagamento = AtualizarDataPagamento(CurrentQuery.FieldByName("PEG").AsInteger, "TRIAGEM2")

                    If (Not vbAtualizouDataPagamento) Then

                          InfoDescription = "Um problema ocorreu ao alterar a data de pagamento do PEG."

                          InserirLogWebService _
                              "RAST", _
                              "TRIAGEM2", _
                              "PEG " + CurrentQuery.FieldByName("PEG").AsString + ": " + InfoDescription

                    End If

                End If

            End If

            SessionVar("ultima_caixa_digitada") = CurrentQuery.FieldByName("ARQUIVOFISICO").AsString
            InfoDescription = "O registro foi incluído."

        Else

            InfoDescription = vsMensagemRetornada

            InserirLogWebService _
                "RAST", _
                "TRIAGEM2", _
                "PEG " + CurrentQuery.FieldByName("PEG").AsString + " - Erro: " + vsMensagemRetornada

        End If

        If ((Not CurrentQuery.FieldByName("PEG").IsNull) And (vbAtualizouDataPagamento)) Then

            Dim qSelect As Object
            Set qSelect = NewQuery

            qSelect.Add("SELECT 1 FROM SAM_PEG WHERE HANDLE = :HANDLE_PEG")
            qSelect.Add("AND SITUACAO = 9 AND NOMEARQUIVOIMPORTADO IS NOT NULL")

            qSelect.ParamByName("HANDLE_PEG").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger

            qSelect.Active = True

            If (Not qSelect.EOF) Then

                ColocarPegEmDigitacao(CurrentQuery.FieldByName("PEG").AsInteger)
                IncluirOcorrenciaPeg(CurrentQuery.FieldByName("PEG").AsInteger)

            End If

            qSelect.Active = False
            Set qSelect = Nothing

        End If

    End If

End Sub

Sub ColocarPegEmDigitacao(piHandlePeg As Long)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    qUpdate.Add("UPDATE SAM_PEG              ")
    qUpdate.Add("   SET SITUACAO = :DIGITACAO")
    qUpdate.Add(" WHERE HANDLE = :HANDLE     ")

    qUpdate.ParamByName("DIGITACAO").AsString  = "1"
    qUpdate.ParamByName("HANDLE"   ).AsInteger = piHandlePeg

    qUpdate.ExecSQL

    Set qUpdate = Nothing

    ColocarGuiasEmDigitacao(piHandlePeg)
    ColocarEventosEmDigitacao(piHandlePeg)

End Sub

Sub ColocarGuiasEmDigitacao(piHandlePeg As Long)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    qUpdate.Add("UPDATE SAM_GUIA              ")
    qUpdate.Add("   SET SITUACAO = :DIGITACAO ")
    qUpdate.Add(" WHERE PEG       = :HANDLEPEG")
    qUpdate.Add("   AND SITUACAO <> :CANCELADA")

    qUpdate.ParamByName("DIGITACAO").AsString  = "1"
    qUpdate.ParamByName("HANDLEPEG").AsInteger = piHandlePeg
    qUpdate.ParamByName("CANCELADA").AsString  = "8"

    qUpdate.ExecSQL

    Set qUpdate = Nothing

End Sub

Sub ColocarEventosEmDigitacao(piHandlePeg As Long)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    qUpdate.Add("UPDATE SAM_GUIA_EVENTOS                      ")
    qUpdate.Add("   SET SITUACAO = :DIGITACAO                 ")
    qUpdate.Add(" WHERE GUIA IN (SELECT HANDLE                ")
    qUpdate.Add("                  FROM SAM_GUIA              ")
    qUpdate.Add("                 WHERE PEG      = :HANDLEPEG ")
    qUpdate.Add("                   AND SITUACAO = :DIGITACAO)")

    qUpdate.ParamByName("DIGITACAO").AsString  = "1"
    qUpdate.ParamByName("HANDLEPEG").AsInteger = piHandlePeg

    qUpdate.ExecSQL

    Set qUpdate = Nothing

End Sub


Sub IncluirOcorrenciaPeg(piHandlePeg As Long)

    Dim qInsert As Object
    Set qInsert = NewQuery

    qInsert.Add("INSERT INTO SAM_PEG_OCORRENCIA (   ")
    qInsert.Add("            HANDLE,                ")
    qInsert.Add("            PEG,                   ")
    qInsert.Add("            DATAHORAINICIAL,       ")
    qInsert.Add("            DATAHORAFINAL,         ")
    qInsert.Add("            USUARIO,               ")
    qInsert.Add("            FASEANTERIOR,          ")
    qInsert.Add("            FASEATUAL,             ")
    qInsert.Add("            OCORRENCIA)            ")
    qInsert.Add("     VALUES (                      ")
    qInsert.Add("            :NOVOHANDLE,           ")
    qInsert.Add("            :HANDLEPEG,            ")
    qInsert.Add("            :AGORA,                ")
    qInsert.Add("            :AGORA,                ")
    qInsert.Add("            :HANDLEUSUARIOCORRENTE,")
    qInsert.Add("            :DEVOLVIDO,            ")
    qInsert.Add("            :DIGITACAO,            ")
    qInsert.Add("            :TEXTO)                ")

    qInsert.ParamByName("NOVOHANDLE"           ).AsInteger  = NewHandle("SAM_PEG_OCORRENCIA")
    qInsert.ParamByName("HANDLEPEG"            ).AsInteger  = piHandlePeg
    qInsert.ParamByName("AGORA"                ).AsDateTime = ServerNow
    qInsert.ParamByName("HANDLEUSUARIOCORRENTE").AsInteger  = CurrentUser
    qInsert.ParamByName("DEVOLVIDO"            ).AsString   = "9"
    qInsert.ParamByName("DIGITACAO"            ).AsString   = "1"
    qInsert.ParamByName("TEXTO"                ).AsString   = "Situação alterada de 'Devolvido' para 'Digitação'."

    qInsert.ExecSQL

    Set qInsert = Nothing

End Sub

Public Sub Atualizar(piHandleRastreamento As Long, piHandlePeg As Long, piHandleCaixa As Long)

    Dim qUpdate As Object
    Set qUpdate = NewQuery

    If (piHandlePeg = Null) Then
        piHandlePeg = 0
    End If

    '-----------------------------------------------------
    'Existia um PEG, mas agora não existe...
    '-----------------------------------------------------
    If ((piHandlePeg >= 0) And (CurrentQuery.FieldByName("PEG").IsNull)) Then
        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG_RASTREADOR          ")
        qUpdate.Add("   SET PEG           = NULL,       ")
        qUpdate.Add("       ARQUIVOFISICO = NULL,       ")
        qUpdate.Add("       DATAEMISSAONF = NULL,       ")
        qUpdate.Add("       NRNOTAFISCAL  = NULL        ")
        qUpdate.Add(" WHERE HANDLE = :HANDLERASTREAMENTO")

        qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = piHandleRastreamento

        qUpdate.ExecSQL

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                            ")
        qUpdate.Add("   SET PEGARQUIVOFISICO = NULL,           ")
        qUpdate.Add("       PEGRASTREADOR    = NULL,           ")
        qUpdate.Add("       DATAEMISSAONOTA  = NULL,           ")
        qUpdate.Add("       NFNUMERO         = NULL            ")
        qUpdate.Add(" WHERE PEGRASTREADOR = :HANDLERASTREAMENTO")

        qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = piHandleRastreamento

        qUpdate.ExecSQL

    End If

    '-----------------------------------------------------
    'Existia um PEG diferente do atual...
    '-----------------------------------------------------
    If ((piHandlePeg > 0) And (piHandlePeg <> CurrentQuery.FieldByName("PEG").AsInteger) And (Not CurrentQuery.FieldByName("PEG").IsNull)) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                            ")
        qUpdate.Add("   SET PEGARQUIVOFISICO = NULL,           ")
        qUpdate.Add("       PEGRASTREADOR    = NULL,           ")
        qUpdate.Add("       DATAEMISSAONOTA  = NULL,           ")
        qUpdate.Add("       NFNUMERO         = NULL            ")
        qUpdate.Add(" WHERE PEGRASTREADOR = :HANDLERASTREAMENTO")
        qUpdate.Add("    OR HANDLE        = :HANDLEPEG         ")

        qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = piHandleRastreamento
        qUpdate.ParamByName("HANDLEPEG"         ).AsInteger = piHandlePeg

        qUpdate.ExecSQL

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                            ")
        qUpdate.Add("   SET PEGRASTREADOR = :HANDLERASTREAMENTO")

        If (Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull) Then

            qUpdate.Add("       ,PEGARQUIVOFISICO = :HANDLEARQUIVOFISICO")
            qUpdate.ParamByName("HANDLEARQUIVOFISICO").AsInteger = CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger

        End If

        qUpdate.Add(" WHERE HANDLE = :HANDLEPEG ")

        qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = piHandleRastreamento
        qUpdate.ParamByName("HANDLEPEG"         ).AsInteger = CurrentQuery.FieldByName("PEG").AsInteger

        qUpdate.ExecSQL

    End If

    '-----------------------------------------------------
    'Não existia um PEG, mas agora existe...
    '-----------------------------------------------------
    If ((piHandlePeg <= 0) And (Not CurrentQuery.FieldByName("PEG").IsNull)) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                            ")
        qUpdate.Add("   SET PEGRASTREADOR = :HANDLERASTREAMENTO")

        If (Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull) Then
            qUpdate.Add("       ,PEGARQUIVOFISICO = :HANDLEARQUIVOFISICO")
            qUpdate.ParamByName("HANDLEARQUIVOFISICO").AsInteger = CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger
        End If

        qUpdate.Add(" WHERE HANDLE = :HANDLEPEG")

        qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = piHandleRastreamento
        qUpdate.ParamByName("HANDLEPEG"         ).AsInteger = CurrentQuery.FieldByName("PEG").AsInteger

        qUpdate.ExecSQL

    End If

    '-----------------------------------------------------
    'Existia uma caixa, mas agora não existe...
    '-----------------------------------------------------
    If ((piHandleCaixa >= 0) And (CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull)) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                            ")
        qUpdate.Add("   SET PEGARQUIVOFISICO = NULL            ")
        qUpdate.Add(" WHERE PEGRASTREADOR = :HANDLERASTREAMENTO")

        qUpdate.ParamByName("HANDLERASTREAMENTO").AsInteger = piHandleRastreamento

        qUpdate.ExecSQL

    End If

    '-----------------------------------------------------
    'Existia uma caixa diferente da atual...
    '-----------------------------------------------------
    If ((piHandleCaixa >= 0) And (piHandleCaixa <> CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger) And (Not CurrentQuery.FieldByName("ARQUIVOFISICO").IsNull)) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                                ")
        qUpdate.Add("   SET PEGARQUIVOFISICO = :HANDLEARQUIVOFISICO")
        qUpdate.Add(" WHERE PEGRASTREADOR = :HANDLERASTREAMENTO    ")

        qUpdate.ParamByName("HANDLEARQUIVOFISICO").AsInteger = CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger
        qUpdate.ParamByName("HANDLERASTREAMENTO" ).AsInteger = piHandleRastreamento

        qUpdate.ExecSQL

    End If

    If ((Not CurrentQuery.FieldByName("DATAEMISSAONF").IsNull) And _
        (Not CurrentQuery.FieldByName("NRNOTAFISCAL").IsNull) And _
        (Not CurrentQuery.FieldByName("QTDGUIAS").IsNull) And _
        (Not CurrentQuery.FieldByName("VLRAPRESENTADO").IsNull) And _
        (Not CurrentQuery.FieldByName("PEG").IsNull)) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG                                  ")
        qUpdate.Add("   SET DATAEMISSAONOTA = :DATAEMISSAONOTAFISCAL,")
        qUpdate.Add("       NFNUMERO        = :NUMERONOTAFISCAL      ")
        qUpdate.Add(" WHERE PEGRASTREADOR = :HANDLERASTREAMENTO      ")

        qUpdate.ParamByName("HANDLERASTREAMENTO"   ).AsInteger  = piHandleRastreamento
        qUpdate.ParamByName("DATAEMISSAONOTAFISCAL").AsDateTime = CurrentQuery.FieldByName("DATAEMISSAONF").AsDateTime
        qUpdate.ParamByName("NUMERONOTAFISCAL"     ).AsInteger  = CurrentQuery.FieldByName("NRNOTAFISCAL").AsString

        qUpdate.ExecSQL

    End If

    If ((CurrentQuery.FieldByName("PEG").IsNull) And _
        (Not CurrentQuery.FieldByName("DATAEMISSAONF").IsNull) And _
        (Not CurrentQuery.FieldByName("NRNOTAFISCAL").IsNull) And _
        (Not CurrentQuery.FieldByName("QTDGUIAS").IsNull) And _
        (Not CurrentQuery.FieldByName("VLRAPRESENTADO").IsNull)) Then

        qUpdate.Clear
        qUpdate.Add("UPDATE SAM_PEG_RASTREADOR                     ")
        qUpdate.Add("   SET DATAEMISSAONF = :DATAEMISSAONOTAFISCAL,")
        qUpdate.Add("       NRNOTAFISCAL  = :NUMERONOTAFISCAL      ")
        qUpdate.Add(" WHERE HANDLE = :HANDLERASTREAMENTO           ")

        qUpdate.ParamByName("HANDLERASTREAMENTO"   ).AsInteger  = piHandleRastreamento
        qUpdate.ParamByName("DATAEMISSAONOTAFISCAL").AsDateTime = CurrentQuery.FieldByName("DATAEMISSAONF").AsDateTime
        qUpdate.ParamByName("NUMERONOTAFISCAL"     ).AsInteger  = CurrentQuery.FieldByName("NRNOTAFISCAL").AsString

        qUpdate.ExecSQL

    End If

    If ((giHandleCaixaAnteEdicao <> CurrentQuery.FieldByName("ARQUIVOFISICO").AsInteger) Or (giHandlePegAnteEdicao <> CurrentQuery.FieldByName("PEG").AsInteger)) Then

        GravarHistoricoRastreamentoCaixa _
            piHandleRastreamento, _
            giHandlePegAnteEdicao, _
            giHandleCaixaAnteEdicao

    End If

    Set qUpdate = Nothing

End Sub
