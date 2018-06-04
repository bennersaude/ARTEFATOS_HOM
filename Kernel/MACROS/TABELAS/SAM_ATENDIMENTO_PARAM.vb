'HASH: 0F8B63331C59D9816BA9337D46BD30AD
'Macro: SAM_ATENDIMENTO_PARAM
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCALCULAR_OnClick()

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT PROCESSADO")
  SQL.Add("  FROM SAM_ATENDIMENTO_PARAM")
  SQL.Add(" WHERE COMPETENCIA = :COMPETENCIA")
  '--------------------------------------------
  SQL.Add("   AND DATAPROCESSAMENTO IS NULL")
  SQL.ParamByName("COMPETENCIA").Value = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime
  SQL.Active = True
  If SQL.FieldByName("PROCESSADO").AsString = "S" Then
    If MsgBox("Competência já calculada. Recalcular?", vbYesNo) = vbNo Then

      Set SQL = Nothing
      Exit Sub
    End If
  End If
  Set SQL = Nothing

  Dim Interface As Object

  If CurrentQuery.FieldByName("PROCESSADO").AsString = "S" Then
    MsgBox("Rotina já Processada.")
    Exit Sub
  End If

  If CurrentQuery.State <>1 Then
    MsgBox("Registro em edição.")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("PROCESSADO").AsString = "N" And CurrentQuery.State = 1 Then
    Set Interface = CreateBennerObject("SamVaga.Atendimento")
    Interface.Calcular(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing
  End If
  RefreshNodesWithTable "SAM_ATENDIMENTO_PARAM"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT PROCESSADO")
  SQL.Add("  FROM SAM_ATENDIMENTO_PARAM")
  '--incluído por Claudemir -01/10/2002-------
  SQL.Add(" WHERE COMPETENCIA = :COMPETENCIA")
  '--------------------------------------------
  SQL.Add(" ORDER BY PROCESSADO")
  SQL.ParamByName("COMPETENCIA").Value = CurrentQuery.FieldByName("COMPETENCIA").AsDateTime

  SQL.Active = True
  If Not SQL.EOF Then
    If SQL.FieldByName("PROCESSADO").AsString = "S" Then
      If bsShowMessage("Competência já cadastrada e calculada. Continuar?", "Q") = vbNo Then
        CanContinue = False
        Set SQL = Nothing
        Exit Sub
      End If
    End If
    If SQL.FieldByName("PROCESSADO").AsString = "N" Then
      bsShowMessage("Competência já cadastrada e não calculada.", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If
  Set SQL = Nothing
End Sub

