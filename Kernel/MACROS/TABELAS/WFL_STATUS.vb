'HASH: 5979876B82DE46F3EA37CDAD04155F9B
'Macro da tabela WFL_STATUS

Option Explicit

'#uses "*AtualizarHistoricoWorkflow"

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

    If (WebVisionCode = "WFL_STATUS_RASTR") Then

        If (CommandID = "K_MOVIMENTAR") Then

            Dim vsMensagem As String
            Dim qSelect    As Object

            Set qSelect = NewQuery

            If (SessionVar("HandleParaMovimentacao") <> "") Then

                qSelect.Clear
                qSelect.Add("SELECT HISTORICOWORKFLOW ")
                qSelect.Add("  FROM SAM_PEG_RASTREADOR")
                qSelect.Add(" WHERE HANDLE = :HANDLE  ")

                qSelect.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HandleParaMovimentacao"))

                qSelect.Active = True

                If (qSelect.EOF) Then
                    CancelDescription = "Rastreamento não localizado."
                    CanContinue = False

                    qSelect.Active = False
                    Set qSelect = Nothing

                    Exit Sub
                End If

                AtualizarHistoricoWorkflow(qSelect.FieldByName("HISTORICOWORKFLOW").AsFloat, _
                                           CLng(SessionVar("HandleParaMovimentacao")), _
                                           CurrentQuery.FieldByName("CODIGO").AsString, _
                                           1, _
                                           1, _
                                           "")

                InfoDescription = "Movimentação para o status '" + CurrentQuery.FieldByName("DESCRICAO").AsString + "' efetuada."

                qSelect.Active = False
                Set qSelect = Nothing

                Exit Sub

            End If

            If (SessionVar("HandleParaMovimentacao_cx") <> "") Then

                Dim vsPegsInertes As String
                Dim vsRetorno     As String
                Dim i             As Integer

                vsPegsInertes = ""
                vsRetorno     = ""

                qSelect.Clear
                qSelect.Add("SELECT HANDLE,                     ")
                qSelect.Add("       HISTORICOWORKFLOW,          ")
                qSelect.Add("       CODRASTREAMENTO             ")
                qSelect.Add("  FROM SAM_PEG_RASTREADOR          ")
                qSelect.Add(" WHERE ARQUIVOFISICO = :HANDLECAIXA")

                qSelect.ParamByName("HANDLECAIXA").AsInteger = CLng(SessionVar("HandleParaMovimentacao_cx"))

                qSelect.Active = True

                i = 0

                While (Not qSelect.EOF)

                    i = i + 1

                    vsRetorno = AtualizarHistoricoWorkflow(qSelect.FieldByName("HISTORICOWORKFLOW").AsInteger, _
                                                           qSelect.FieldByName("HANDLE").AsInteger, _
                                                           CurrentQuery.FieldByName("CODIGO").AsString, _
                                                           1, _
                                                           1, _
                                                           "")

                    If (vsRetorno <> "OK") Then
                        i = i - 1
                        vsPegsInertes = vsPegsInertes + qSelect.FieldByName("CODRASTREAMENTO").AsString + "->" + vsRetorno + vbCrLf
                    End If

                    vsRetorno = "OK"

                    qSelect.Next

                Wend

                qSelect.Active = False
                Set qSelect = Nothing

                If (vsPegsInertes = "") And (vsRetorno = "") Then

                    InfoDescription = "A caixa informada não foi localizada."

                ElseIf (vsPegsInertes = "") And (vsRetorno = "OK") Then

                    InfoDescription = "A caixa com " + CStr(i) + " PEGs teve o status alterado para '" + CurrentQuery.FieldByName("DESCRICAO").AsString + "'."

                Else

                    If (i = 0) Then

                        InfoDescription = "Nenhum PEG foi movimentado."

                    Else

                        If (i = 1) Then

                            InfoDescription = "1 PEG movimentado para '" + CurrentQuery.FieldByName("DESCRICAO").AsString + "'.

                        Else

                            InfoDescription = CStr(i) + " PEGs movimentados para '" + CurrentQuery.FieldByName("DESCRICAO").AsString + "'."

                        End If

                        InfoDescription = InfoDescription + " Os PEGs abaixo não foram movimentados:" + vbCrLf + vsPegsInertes

                    End If

                End If

                Exit Sub

            End If

        End If

    End If

End Sub
