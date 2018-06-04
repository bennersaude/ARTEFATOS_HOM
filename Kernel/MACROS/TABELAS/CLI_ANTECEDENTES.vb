'HASH: C94A81314596D4A9EB7DC06911E977AA
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  If Not AtendimentoEstaAberto Then
    CanContinue = False
    bsShowMessage("Para excluir um registro é necessário que o atendimento esteja aberto!","E")
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If Not AtendimentoEstaAberto Then
    CanContinue = False
    bsShowMessage("Para alterar um registro é necessário que o atendimento esteja aberto!","E")
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If CLng(SessionVar("SAM_MATRICULA")) <= 0 Then
    CanContinue = False
   bsShowMessage("É necessário selecionar um paciente e um atendimento!","E")
  End If

  If CLng(SessionVar("CLI_SUBJETIVO")) <= 0 Then
    CanContinue = False
    bsShowMessage("É necessário selecionar um atendimento!","E")
  End If

  If Not AtendimentoEstaAberto Then
    CanContinue = False
    bsShowMessage("Para incluir um novo registro é necessário que o atendimento esteja aberto!","E")
  End If

End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("ATENDIMENTO").AsInteger = CLng(SessionVar("CLI_SUBJETIVO"))
  CurrentQuery.FieldByName("MATRICULA").AsInteger   = CLng(SessionVar("SAM_MATRICULA"))
End Sub

Public Function AtendimentoEstaAberto As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT DATAENCERRAMENTO")
  SQL.Add("  FROM CLI_SUBJETIVO   ")
  SQL.Add(" WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CLng(SessionVar("CLI_SUBJETIVO"))
  SQL.Active = True

  If SQL.FieldByName("DATAENCERRAMENTO").IsNull Then
    AtendimentoEstaAberto = True
  Else
    AtendimentoEstaAberto = False
  End If

  SQL.Active = False
  Set SQL = Nothing
End Function

Public Function ExisteHabitosDeVida As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear

  SQL.Add("SELECT 1                           ")
  SQL.Add("  FROM CLI_ANTECEDENTES            ")
  SQL.Add(" WHERE ATENDIMENTO = :HATENDIMENTO ")
  SQL.Add("   AND TABTIPO = 2                 ")
  SQL.ParamByName("HATENDIMENTO").AsInteger = CurrentQuery.FieldByName("ATENDIMENTO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    ExisteHabitosDeVida = True
  Else
    ExisteHabitosDeVida = False
  End If

  SQL.Active = False
  Set SQL = Nothing

End Function

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If (ExisteHabitosDeVida) And (CurrentQuery.FieldByName("TABTIPO").AsInteger = 2) And (CurrentQuery.State = 3) Then
    CanContinue = False
    bsShowMessage("Já existe registro de hábitos de vida para esse atendimento!","E")
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 1 Then
    If TipoMedicamento And CurrentQuery.FieldByName("DADOSMEDICAMENTO").IsNull Then
      CanContinue = False
      bsShowMessage("Para alergias do tipo medicamento é necessário que se informe o campo 'Dados medicamento'!","E")
    End If
  End If

End Sub

Public Function TipoMedicamento As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT MEDICAMENTO           ")
  SQL.Add("  FROM CLI_TIPOALERGIA       ")
  SQL.Add(" WHERE HANDLE = :TIPOALERGIA ")
  SQL.ParamByName("TIPOALERGIA").AsInteger = CurrentQuery.FieldByName("TIPOALERGIA").AsInteger
  SQL.Active = True
  If SQL.FieldByName("MEDICAMENTO").AsString = "S" Then
    TipoMedicamento = True
  Else
    TipoMedicamento = False
  End If
End Function
