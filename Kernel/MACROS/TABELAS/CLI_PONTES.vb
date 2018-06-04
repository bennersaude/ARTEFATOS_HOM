'HASH: D9FE4B5C7B374E4C559DD9E5C0EE00A3
'#Uses "*bsShowMessage"
'CLI_PONTES
Option Explicit
Public Sub BOTAOPROCESSAR_OnClick()
  Dim SQL As Object
  Dim Interface As Object
  Set Interface = CreateBennerObject("BSCli012.rotinas")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM CLI_PONTESRECURSO WHERE PONTE = :PONTE")
  SQL.ParamByName("PONTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If

  Set SQL = Nothing
  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterEdit()
  BOTAOPROCESSAR.Enabled = False
End Sub

Public Sub TABLE_AfterInsert()
  BOTAOPROCESSAR.Enabled = False
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM CLI_PONTESRECURSO WHERE PONTE = :PONTE")
  SQL.ParamByName("PONTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM CLI_PONTESRECURSO WHERE PONTE = :PONTE")
  SQL.ParamByName("PONTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True
  If Not SQL.EOF Then
    BOTAOPROCESSAR.Enabled = False
  Else
    BOTAOPROCESSAR.Enabled = True
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_OnDeleteBtnClick(CanContinue As Boolean)
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição ou inserção!", "I")
    Exit Sub
  End If

  Dim Interface As Object
  Set Interface =CreateBennerObject("BSCli012.rotinas")

  If Interface.PodeExcluirExcecao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) Then

    Dim qExcluir As Object
    Set qExcluir = NewQuery

    qExcluir.Active = False
    qExcluir.Clear
    qExcluir.Add("DELETE                                           ")
    qExcluir.Add("  FROM CLI_PONTESRECURSOAGENDA                   ")
    qExcluir.Add(" WHERE PONTESRECURSO IN (SELECT HANDLE           ")
    qExcluir.Add("                           FROM CLI_PONTESRECURSO")
    qExcluir.Add("                          WHERE PONTE = :PONTE)  ")
    qExcluir.ParamByName("PONTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qExcluir.ExecSQL

    qExcluir.Active = False
    qExcluir.Clear
    qExcluir.Add("DELETE                  ")
    qExcluir.Add("  FROM CLI_PONTESRECURSO")
    qExcluir.Add(" WHERE PONTE = :PONTE   ")
    qExcluir.ParamByName("PONTE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qExcluir.ExecSQL

    Set qExcluir = Nothing

    CanContinue = True

  Else
    CanContinue = False
  End If

  Set Interface = Nothing
End Sub
