'HASH: 89B30EE1D0BC260C8C41D0C7D768B477
'Macro : POS_LOTE

Public Sub BOTAOCORRECAO_OnClick()
  Dim SQL As Object
  Dim vLog As String
  Dim vUsuario As String
  Set SQL = NewQuery

  SQL.Add("SELECT NOME ")
  SQL.Add("  FROM Z_GRUPOUSUARIOS ")
  SQL.Add(" WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("USUARIO").AsInteger
  SQL.Active = True

  vUsuario = SQL.FieldByName("NOME").AsString
  vLog = ""
  vLog = Chr(13) + _
         "-------------------------------------------------------------------------------------" + Chr(10) + _
         "      Lote com falha de execução alterado manualmente.                               " + Chr(10) + _
         "      Pelo usuario          : " + vUsuario + Chr(10) + _
         "      Data                  : " + Str(ServerNow) + Chr(10) + _
         "-------------------------------------------------------------------------------------"

  SQL.Active = False
  SQL.Clear
  If MsgBox("Atenção!Para mudar o status do lote verifique se não tem processo em andamento.Tem certeza que deseja continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    SQL.Add("UPDATE POS_LOTE")
    SQL.Add("   SET TABSITUACAO = 2,         ")
    SQL.Add("       RESULTADOPROCESSO = 'E',         ")
    If InStr(SQLServer, "MSSQL") > 0 Then
      SQL.Add("       OBSERVACOES  = SUBSTRING(OBSERVACOES,1,4000) + '" + vLog + "'")
    Else
      SQL.Add("       OBSERVACOES  = OBSERVACOES || '" + vLog + "'")
    End If
    SQL.Add(" WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL
  End If

  RefreshNodesWithTable("POS_LOTE")

End Sub

Public Sub BOTAOPROCESSARLOTE_OnClick()

  '*********************************************************************************
  '******************** ALTERAÇÃO PARA POS - PROCESSANDO LOTES *********************
  '*********************************************************************************
  Dim interface As Object
  Dim SQL As Object
  Dim vTipo As String
  Dim vHandle As Long
  'MsgBox  CurrentQuery.State
  If CurrentQuery.State = 1 Then
    If CurrentQuery.FieldByName("TIPOLOTE").AsString = "A" Then
      Set SQL = NewQuery
      SQL.Add("SELECT TABSITUACAO FROM POS_LOTE WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.Active = True
      If SQL.FieldByName("TABSITUACAO").AsInteger = 1 Then
        vHandle = CurrentQuery.FieldByName("HANDLE").AsInteger
        vTipo = CurrentQuery.FieldByName("TIPOLOTE").AsString
        Set interface = CreateBennerObject("BSATE003.Rotinas")
        interface.ExecAutorizUsuario(CurrentSystem, vHandle, vTipo)
        Set interface = Nothing
        RefreshNodesWithTable("POS_LOTE")
      Else
        MsgBox("Lote já processado.")
        Exit Sub
      End If
    Else
      Set SQL = NewQuery
      SQL.Add("SELECT TABSITUACAO FROM POS_LOTE WHERE HANDLE = :HANDLE")
      SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.Active = True
      If SQL.FieldByName("TABSITUACAO").AsInteger = 1 Then
        vHandle = CurrentQuery.FieldByName("HANDLE").AsInteger
        vTipo = CurrentQuery.FieldByName("TIPOLOTE").AsString
        Set interface = CreateBennerObject("BSATE003.Rotinas")
        interface.ExecAutorizUsuario(CurrentSystem, vHandle, vTipo)
        Set interface = Nothing
        RefreshNodesWithTable("POS_LOTE")
      Else
        MsgBox("Lote já processado.")
        Exit Sub
      End If
    End If
  Else
    MsgBox("O registro não pode estar em edição.")
    Exit Sub
  End If

  '******************* FIM DA ALTERAÇÃO ********************************************
  '    SQL.AQCTIVE = False

End Sub

Public Sub BOTAOREPROCESSARLOTES_OnClick()
  Dim interface As Object
  Dim vHandle As Long
  If CurrentQuery.State = 1 Then
    vHandle = CurrentQuery.FieldByName("HANDLE").AsInteger
    Set interface = CreateBennerObject("BSATE003.Rotinas")
    interface.Reprocessar(CurrentSystem, vHandle)
    Set interface = Nothing
    RefreshNodesWithTable("POS_LOTE")
  Else
    MsgBox("O registro não pode estar em edição.")
    Exit Sub
  End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Dim vTipo As String
  Dim vHandle As Long
  Set SQL = NewQuery
  SQL.Add("SELECT COUNT(1) QTDREG FROM POS_LOTE WHERE TABSITUACAO = 1 AND HANDLE <> :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If SQL.FieldByName("QTDREG").AsInteger > 0 Then
    MsgBox("Só é permitido um registro aberto de cada vez.")
    CanContinue = False
    Exit Sub
  End If
  CurrentQuery.FieldByName("USUARIO").Value = CurrentUser
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("DATA").Value = ServerNow
End Sub

