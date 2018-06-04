'HASH: C4F7C3BE79B0683D91E99905B2818D2B
'Macro: SAM_GUIA_DEVOLUCAO_GLOSA
Option Explicit


Public Sub BOTAOGLOSA_OnClick()
  Dim INTERFACE As Object
  Set INTERFACE = CreateBennerObject("samPEG.PROCESSAR")
  INTERFACE.INICIALIZAR(CurrentSystem)

  INTERFACE.DESCRICAOGLOSA(CurrentSystem, CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger)
  INTERFACE.FINALIZAR
  Set INTERFACE = Nothing
End Sub

Public Sub MOTIVOGLOSA_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim INTERFACE As Object
  ShowPopup = False

  vColunas = "CODIGOGLOSA|SAM_MOTIVOGLOSA.DESCRICAO|SAM_MOTIVOGLOSA.EXIGECOMPLEMENTO|SAM_TIPOMOTIVOGLOSA.DESCRICAO"
  vCampos = "Código|Descrição|Complemento|Tipo Glosa"
  vCriterio = "SAM_MOTIVOGLOSA.ATIVA='S'"

  Set INTERFACE = CreateBennerObject("Procura.Procurar")
  vHandle = INTERFACE.Exec(CurrentSystem, "SAM_MOTIVOGLOSA|SAM_TIPOMOTIVOGLOSA[SAM_MOTIVOGLOSA.TIPOMOTIVOGLOSA=SAM_TIPOMOTIVOGLOSA.HANDLE]", vColunas, 2, vCampos, vCriterio, "Tabela de Motivo de Glosas", True, "")
  Set INTERFACE = Nothing

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("MOTIVOGLOSA").Value = vHandle
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Cancontinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "E")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTable("SAM_GUIA_DEVOLUCAO_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Cancontinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "A")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTable("SAM_GUIA_DEVOLUCAO_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Cancontinue = CheckFilialProcessamento(CurrentSystem, RecordHandleOfTable("FILIAIS"), "I")
  'Somente dara manutencao se for sua filial ou se sua filial for a de processamento
  If CanContinue = False Then
    RefreshNodesWithTable("SAM_GUIA_DEVOLUCAO_GLOSA")'Colocar tabela para refresh
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT ATIVA FROM SAM_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  sql.Active = True
  If sql.FieldByName("ATIVA").AsString = "N" Then
    MsgBox("O motivo de glosa está desativado")
    CanContinue = False
    Exit Sub
  End If


  Dim vExigeComplemento As String
  sql.Clear
  SQL.Add("SELECT EXIGECOMPLEMENTO FROM SAM_MOTIVOGLOSA WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").AsInteger
  SQL.Active = True
  vExigeComplemento = SQL.FieldByName("EXIGECOMPLEMENTO").AsString

  If vExigeComplemento = "S" Then
    If Len(CurrentQuery.FieldByName("COMPLEMENTO").AsString)<4 Then
      MsgBox "Motivo de glosa exige a digitação do complemento!"
      Cancontinue = False
      ' COMPLEMENTO.SetFocus
    End If

  End If
  SQL.Active = False
  Set SQL = Nothing

End Sub

