'HASH: A574CFE0B7247F587C1692C533D4AB7B
'Macro da tabela SAM_DEPARADOR_PRESTADORES

Option Explicit

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

    If ((WebVisionCode = "DEPARADORPRESTADORES") Or (WebVisionCode = "DEPARADORITENS")) Then

        If (CurrentQuery.FieldByName("USUARIORESPONSAVEL").AsInteger <> CurrentUser) Then
            CancelDescription = "Somente o resposável pela demanda pode editá-la."
            CanContinue = False
            Exit Sub
        End If

        If (CurrentQuery.FieldByName("QUANTIDADEITENS").AsInteger = 0) Then
            CancelDescription = "A análise já foi concluída."
            CanContinue = False
            Exit Sub
        End If
    End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

    Select Case CommandID

        Case "K_ASSUMIR"

            If (CurrentQuery.FieldByName("USUARIORESPONSAVEL").IsNull) Then
                AtualizarUsuarioPara(CurrentUser)
            Else
                If (RequestConfirmation("Deseja assumir?")) Then
                    AtualizarUsuarioPara(CurrentUser)
                End If
            End If

        Case "K_DEVOLVER"

            If (CurrentQuery.FieldByName("QUANTIDADEITENS").AsInteger = 0) Then
                CancelDescription = "A análise já foi concluída."
                CanContinue = False
                Exit Sub
            End If

            AtualizarUsuarioPara(0)

    End Select

End Sub

Public Sub AtualizarUsuarioPara(piHandle As Long)

    CurrentQuery.Edit

    If (piHandle > 0) Then
        CurrentQuery.FieldByName("USUARIORESPONSAVEL").AsInteger = piHandle
    Else
        CurrentQuery.FieldByName("USUARIORESPONSAVEL").Clear
    End If

    CurrentQuery.Post

End Sub
