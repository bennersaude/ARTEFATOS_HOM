'HASH: DA6DEC5F4CF32319291B3441E7A63EF6
'#Uses "*bsShowMessage"

'CLI_ROTREMARCAFALTA

Public Sub BOTAOPROCESSAR_OnClick()
  If CurrentQuery.State <>1 Then
    MsgBox("O registro não pode estar em edição ou inserção!")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    Dim BSCLI001DLL As Object
    Set BSCLI001DLL = CreateBennerObject("BSCLI001.ROTINAS")
    BSCLI001DLL.BuscaFaltaRecurso(CurrentSystem, _
                                  CurrentQuery.FieldByName("DATA").AsDateTime, _
                                  CurrentQuery.FieldByName("CLINICA").AsInteger, _
                                  CurrentQuery.FieldByName("RECURSO").AsInteger, _
                                  CurrentQuery.FieldByName("ESCALA").AsInteger, _
                                  CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set BSCLI001DLL = Nothing
    RefreshNodesWithTable("CLI_ROTREMARCAFALTA")
  Else
    MsgBox("Rotina já foi processada!")
  End If
End Sub

Public Sub ESCALA_OnPopup(ShowPopup As Boolean)
  ESCALA.LocalWhere = "CLI_ESCALA.DISPONIVEL = 'S'"
End Sub

Public Sub TABLE_AfterScroll()
  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    MsgBox("Rotina já foi processada!")
    CanContinue = False
    Exit Sub
  End If

  Dim sql As BPesquisa
  Set sql = NewQuery
  sql.Clear

  sql.Add("SELECT COUNT(1) QTDE ")
  sql.Add("FROM CLI_ROTREMARCAFALTA ")
  sql.Add("WHERE RECURSO = :RECURSO ")
  sql.Add("AND DATA = :DATA ")
  sql.Add("AND DATAPROCESSAMENTO IS NOT NULL ")

  sql.ParamByName("RECURSO").Value = CurrentQuery.FieldByName("RECURSO").Value
  sql.ParamByName("DATA").Value = CurrentQuery.FieldByName("DATA").Value

  sql.Active = True

  If (sql.FieldByName("QTDE").AsInteger > 0) Then
    bsShowMessage("Essa data já foi processada para o recurso selecionado.", "I")
    Set sql = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set sql = Nothing

End Sub

