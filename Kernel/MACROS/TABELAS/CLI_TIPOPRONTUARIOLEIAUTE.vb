'HASH: 7FD2FD548BF33CEFD7FB92D9B3D33C78
'CLI_TIPOPRONTUARIOLEIAUTE
'SHIBA
Dim CAMPO As Integer

Public Sub TABLE_AfterInsert()
End Sub

Public Sub TABLE_AfterPost()
  Dim SQL As Object
  Set SQL = NewQuery
  Dim INSERT As Object
  Set INSERT = NewQuery
  Dim VERIFICA As Object
  Set VERIFICA = NewQuery
  VERIFICA.Clear
  VERIFICA.Add("SELECT 1 FROM CLI_PRONTUARIOCONTEUDO")
  VERIFICA.Add("WHERE PRONTUARIO = :PRONTUARIO AND TIPOPRONTUARIOLEIAUTE = :TIPOPRONTUARIOLEIAUTE")
  INSERT.Clear
  INSERT.Add("INSERT INTO CLI_PRONTUARIOCONTEUDO")
  INSERT.Add("(HANDLE, TIPOPRONTUARIOLEIAUTE, ORDEM, PRONTUARIO)")
  INSERT.Add("VALUES")
  INSERT.Add("(:HANDLE, :TIPOPRONTUARIOLEIAUTE, :ORDEM, :PRONTUARIO)")
  SQL.Clear
  SQL.Add("SELECT DISTINCT PRONTUARIO FROM CLI_PRONTUARIOCONTEUDO")
  SQL.Add("WHERE TIPOPRONTUARIOLEIAUTE IN (SELECT HANDLE FROM CLI_TIPOPRONTUARIOLEIAUTE WHERE TIPOPRONTUARIO = :TIPOPRONTUARIO)")
  SQL.ParamByName("TIPOPRONTUARIO").AsInteger = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
  SQL.Active = True
  While Not SQL.EOF
    VERIFICA.Active = False
    VERIFICA.ParamByName("PRONTUARIO").AsInteger = SQL.FieldByName("PRONTUARIO").AsInteger
    VERIFICA.ParamByName("TIPOPRONTUARIOLEIAUTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    VERIFICA.Active = True
    If VERIFICA.EOF Then
      INSERT.ParamByName("HANDLE").AsInteger = NewHandle("CLI_PRONTUARIOCONTEUDO")
      INSERT.ParamByName("TIPOPRONTUARIOLEIAUTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      INSERT.ParamByName("ORDEM").AsInteger = CurrentQuery.FieldByName("ORDEM").AsInteger
      INSERT.ParamByName("PRONTUARIO").AsInteger = SQL.FieldByName("PRONTUARIO").AsInteger
      INSERT.ExecSQL
    End If
    SQL.Next
  Wend
  Set SQL = Nothing
  Set INSERT = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  CAMPO = CurrentQuery.FieldByName("TABTIPOCAMPO").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim LEIAUTE As Object
  Dim SQL As Object
  Dim ValidacaoCampo As Object
  Dim vOk As Boolean

  If(CAMPO <>CurrentQuery.FieldByName("TABTIPOCAMPO").AsInteger)And(CurrentQuery.State <>3)Then
  CanContinue = False
  MsgBox("Não é possível trocar o tipo do campo!")
  Exit Sub
End If

Set LEIAUTE = NewQuery
If Not InTransaction Then StartTransaction

LEIAUTE.Add("UPDATE CLI_PRONTUARIOCONTEUDO                         ")
LEIAUTE.Add("   SET ORDEM = :ORDEM                                 ")
LEIAUTE.Add(" WHERE TIPOPRONTUARIOLEIAUTE = :TIPOPRONTUARIOLEIAUTE ")
LEIAUTE.ParamByName("ORDEM").Value = CurrentQuery.FieldByName("ORDEM").AsInteger
LEIAUTE.ParamByName("TIPOPRONTUARIOLEIAUTE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
LEIAUTE.ExecSQL

If InTransaction Then Commit
Set LEIAUTE = Nothing

Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT NOMECAMPO")
SQL.Add("  FROM CLI_TIPOPRONTUARIOLEIAUTE")
SQL.Add(" WHERE NOMECAMPO =:NOME")
SQL.Add("   AND HANDLE<>:HANDLE")
SQL.Add("   AND TIPOPRONTUARIO = :TIPOPRONTUARIO")
SQL.ParamByName("NOME").AsString = CurrentQuery.FieldByName("NOMECAMPO").AsString
SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
SQL.ParamByName("TIPOPRONTUARIO").AsInteger = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
SQL.Active = True

If Not(SQL.FieldByName("NOMECAMPO").IsNull)Then
  MsgBox("Nome do Campo Já Existente!")
  CanContinue = False
  Exit Sub
End If

If CurrentQuery.FieldByName("TABTIPOCAMPO").AsInteger = 6 Then
  SQL.Clear
  SQL.Add("SELECT COUNT(*) TOTAL")
  SQL.Add("  FROM CLI_TIPOPRONTUARIOLEIAUTE")
  SQL.Add(" WHERE TABTIPOCAMPO = 6")
  SQL.Add("   AND HANDLE<>:HANDLE")
  SQL.Add("   AND TIPOPRONTUARIO = :TIPOPRONTUARIO")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("TIPOPRONTUARIO").AsInteger = CurrentQuery.FieldByName("TIPOPRONTUARIO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TOTAL").AsInteger >0 Then
    MsgBox("Já existe um campo do tipo memo!")
    CanContinue = False
    Exit Sub
  End If
End If

Set ValidacaoCampo = CreateBennerObject("Cliprontuario.Rotinas")
ValidacaoCampo.ValidaCampo(CurrentSystem, CurrentQuery.FieldByName("NOMECAMPO").AsString, vOk)
Set ValidacaoCampo = Nothing
If vOk = False Then
  MsgBox("Nome do Campo Inválido, possui espaços ou caracteres não permitidos!")
  CanContinue = False
  Exit Sub
End If

If CurrentQuery.FieldByName("TABTIPOCAMPO").AsInteger = 9 Then
  Dim VALIDAR As Object
  Set VALIDAR = CreateBennerObject("CliProntuario.Rotinas")
  If Not VALIDAR.VALIDAITEMS(CurrentSystem, CurrentQuery.FieldByName("ITEMS").AsString)Then
    MsgBox("Campo Itens inválido!")
    CanContinue = False
    Exit Sub
  End If
  Set VALIDAR = Nothing
End If

End Sub

