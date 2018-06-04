'HASH: 3F02CBBFB3DC8D6271C0B8B69EDACA73
 
Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
    If (CommandID = "SUSPENDER") Then
        InfoDescription = suspender()
    End If
End Sub

Public Function suspender As String
    On Error GoTo erro
    CurrentQuery.Edit
    CurrentQuery.FieldByName("DATAFINAL").AsDateTime = CurrentVirtualQuery.FieldByName("DATAFINAL").AsDateTime
    CurrentQuery.Post
    suspender = ""
    Exit Function
erro:
    suspender = Err.Description
End Function

Public Sub TABLE_AfterPost()
  If (CurrentQuery.FieldByName("USOCONTINUO").AsString = "S") Then
    InserirUsoContinuo
  End If
End Sub

Public Sub InserirUsoContinuo
    Dim sql As BPesquisa
    Set sql=NewQuery
    sql.Add("SELECT HANDLE FROM CLI_MEDICAMENTOUSOCONTINUO WHERE PRESCRICAOITEM=:P")
    sql.ParamByName("P").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.Active=True

    Dim handle As Long
    handle = sql.FieldByName("HANDLE").AsInteger
    If (handle = 0) Then
        ' captura a matricula do paciente
        Dim matricula As Long

        sql.Active = False
        sql.Clear
        sql.Add("SELECT S.PACIENTE                                         ")
        sql.Add("  FROM CLI_SUBJETIVO        S                             ")
        sql.Add("  JOIN CLI_PRESCRICAO       P ON P.ATENDIMENTO = S.HANDLE ")
        sql.Add("  JOIN CLI_PRESCRICAO_ITEM PI ON PI.PRESCRICAO = P.HANDLE ")
        sql.Add(" WHERE PI.HANDLE = :P                                     ")
        sql.Active = True

        matricula = sql.FieldByName("PACIENTE").AsInteger

        ' tentar localizar o medicamento de uso continuo igual ao item da prescricao e que esteja ativo
        sql.Active = False
        sql.Clear
        sql.Add("SELECT A.HANDLE                                                                         ")
        sql.Add("  FROM CLI_MEDICAMENTOUSOCONTINUO A                                                     ")
        sql.Add(" WHERE A.MEDICAMENTO=:C                                                                 ")
        sql.Add("   AND (A.DATAFINAL IS NULL OR A.DATAFINAL >=:D)                                        ")
        sql.Add("   AND A.PRESCRICAO IN (SELECT P.HANDLE                                                 ")
        sql.Add("                            FROM CLI_PRESCRICAO P                                       ")
        sql.Add("                            JOIN CLI_SUBJETIVO  S ON S.HANDLE = P.ATENDIMENTO           ")
        sql.Add("                           WHERE S.PACIENTE = :M)                                       ")
        sql.ParamByName("C").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger
        sql.ParamByName("D").AsDateTime = ServerDate
        sql.ParamByName("M").AsInteger = matricula
        sql.Active=True
        handle = sql.FieldByName("HANDLE").AsInteger
        If (handle > 0) Then
            'VINCULAR O MEDICAMENTO DE USO CONTINUO NA PRESCRICAO
            sql.Active = False
            sql.Clear
            sql.Add("UPDATE CLI_MEDICAMENTOUSOCONTINUO SET PRESCRICAOITEM=:PRESCRICAOITEM, PRESCRICAO = :PRESCRICAO WHERE HANDLE=:HANDLE")

            sql.ParamByName("HANDLE").AsInteger         = handle
            sql.ParamByName("PRESCRICAOITEM").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
            sql.ParamByName("PRESCRICAO").AsInteger     = CurrentQuery.FieldByName("PRESCRICAO").AsInteger
            sql.ExecSQL
        Else
          'INSERIR
          sql.Active = False
          sql.Clear
          sql.Add("INSERT INTO CLI_MEDICAMENTOUSOCONTINUO (HANDLE, MEDICAMENTO, DATAINICIAL, INSTRUCOES, PRESCRICAOITEM, PRESCRICAO) VALUES ")
          sql.Add("(:HANDLE, :MEDICAMENTO, :DATAINICIAL, :INSTRUCOES, :PRESCRICAOITEM, :PRESCRICAO)")

          handle = NewHandle("CLI_MEDICAMENTOUSOCONTINUO")
          sql.ParamByName("HANDLE").AsInteger = handle
          sql.ParamByName("MEDICAMENTO").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger
          sql.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
          sql.ParamByName("INSTRUCOES").AsString = CurrentQuery.FieldByName("INSTRUCOES").AsString
          sql.ParamByName("PRESCRICAOITEM").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          sql.ParamByName("PRESCRICAO").AsInteger = CurrentQuery.FieldByName("PRESCRICAO").AsInteger
          sql.ExecSQL
        End If
    Else
        'ATUALIZAR O MEDICAMENTO CONFORME A PRESCRICAO
        sql.Active = False
        sql.Clear
        sql.Add("UPDATE CLI_MEDICAMENTOUSOCONTINUO SET MEDICAMENTO=:MEDICAMENTO, DATAINICIAL=:DATAINICIAL, ")
        sql.Add("INSTRUCOES=:INSTRUCOES, PRESCRICAOITEM=:PRESCRICAOITEM WHERE HANDLE=:HANDLE")

        sql.ParamByName("HANDLE").AsInteger = handle
        sql.ParamByName("MEDICAMENTO").AsInteger = CurrentQuery.FieldByName("ITEM").AsInteger
        sql.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
        sql.ParamByName("INSTRUCOES").AsString = CurrentQuery.FieldByName("INSTRUCOES").AsString
        sql.ParamByName("PRESCRICAOITEM").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        sql.ExecSQL
    End If

    Set sql=Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  ExcluirUsoContinuo
  ExcluirCID
End Sub

Public Sub ExcluirCID
    Dim sql As BPesquisa
    Set sql=NewQuery
    sql.Add("DELETE FROM CLI_PRESCRICAO_ITEM_CID WHERE PRESCRICAOITEM=:I")
    sql.ParamByName("I").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ExecSQL
    Set sql=Nothing
End Sub

Public Sub ExcluirUsoContinuo
    Dim sql As BPesquisa
    Set sql=NewQuery
    sql.Add("DELETE FROM CLI_MEDICAMENTOUSOCONTINUO_CID WHERE MEDICAMENTOUSOCONTINUO IN (SELECT HANDLE FROM CLI_MEDICAMENTOUSOCONTINUO WHERE PRESCRICAOITEM=:I)")
    sql.ParamByName("I").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ExecSQL

    sql.Clear
    sql.Add("DELETE FROM CLI_MEDICAMENTOUSOCONTINUO WHERE PRESCRICAOITEM=:I")
    sql.ParamByName("I").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.ExecSQL

    Set sql=Nothing
End Sub


