'HASH: 9CB4FD60FD7813A93741B5D2E9B0A11B
'#Uses "*CriaTabelaTemporariaSqlServer"
Option Explicit

Public Sub TABLE_AfterInsert()
  Dim vPeg   As Long
  Dim vGuia  As Long
  Dim vChave As Long

  vPeg  = CLng(SessionVar("MONITORAMENTOPEG"))
  vGuia = CLng(SessionVar("MONITORAMENTOGUIA"))

  If vPeg > 0 Then
    vChave = vPeg
  Else
    vChave = vGuia
  End If

  If Not InTransaction Then
    StartTransaction
  End If

  On Error GoTo Erro

    Dim PrefixoTmp As String
    If InStr(SQLServer, "MSSQL") > 0 Then
      CriaTabelaTemporariaSqlServer
      PrefixoTmp = "#"
    ElseIf InStr(SQLServer, "ORACLE") > 0 Then
      PrefixoTmp = ""
    End If

    Dim spProc As BStoredProc
    Set spProc = NewStoredProc
    spProc.Name = "BS_B7F1F4B7"
    spProc.AddParam("P_CHAVE", ptInput, ftInteger)
    spProc.AddParam("P_HANDLEPEG", ptInput, ftInteger)
    spProc.AddParam("P_HANDLEGUIA", ptInput, ftInteger)
    spProc.AddParam("P_ORIGEMPROCESSO", ptInput, ftString)
    spProc.AddParam("P_EXISTEPENDENCIA", ptOutput, ftString)
    spProc.ParamByName("P_CHAVE").AsInteger = vChave
    spProc.ParamByName("P_HANDLEPEG").AsInteger = vPeg
    spProc.ParamByName("P_HANDLEGUIA").AsInteger = vGuia
    spProc.ParamByName("P_ORIGEMPROCESSO").AsString = "M"
    spProc.ExecProc
    Set spProc = Nothing

    Dim vMensagem As String
    vMensagem = ""

    Dim qBuscaMensagemMonitoramento As Object
    Set qBuscaMensagemMonitoramento = NewQuery
    qBuscaMensagemMonitoramento.Clear
    qBuscaMensagemMonitoramento.Add("SELECT TEXTO                          ")
    qBuscaMensagemMonitoramento.Add("  FROM " + PrefixoTmp + "TMP_MENSAGEM ")
    qBuscaMensagemMonitoramento.Add(" WHERE CHAVE = :CHAVE                 ")
    qBuscaMensagemMonitoramento.Add(" ORDER BY ORDEM                       ")
    qBuscaMensagemMonitoramento.ParamByName("CHAVE").AsInteger = vChave
    qBuscaMensagemMonitoramento.Active = True

    If Not qBuscaMensagemMonitoramento.EOF Then
      While Not qBuscaMensagemMonitoramento.EOF
        vMensagem = vMensagem + qBuscaMensagemMonitoramento.FieldByName("TEXTO").AsString
        qBuscaMensagemMonitoramento.Next
      Wend
    Else
      vMensagem = "Não existe guias pendentes para o envio do Monitoramento!"
    End If

    qBuscaMensagemMonitoramento.Clear
    qBuscaMensagemMonitoramento.Add("DELETE FROM " + PrefixoTmp + "TMP_MENSAGEM WHERE CHAVE = :CHAVE")
    qBuscaMensagemMonitoramento.ParamByName("CHAVE").AsInteger = vChave
    qBuscaMensagemMonitoramento.ExecSQL

    Set qBuscaMensagemMonitoramento = Nothing
    If InTransaction Then
      Commit
    End If

    CurrentQuery.FieldByName("RESUMO").AsString = vMensagem
    Exit Sub

  Erro:
    CurrentQuery.FieldByName("RESUMO").AsString = Err.Description
    If InTransaction Then
      Rollback
    End If
End Sub
