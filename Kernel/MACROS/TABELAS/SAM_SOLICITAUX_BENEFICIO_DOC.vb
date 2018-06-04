'HASH: C54259BDD64103BBD5EB44B0A5DC1576
'SAM_SOLICITAUX_BENEFICIO_DOC

Option Explicit

Public Sub TABLE_AfterInsert()

  Exit Sub

  Dim SQLDOC As Object
  Dim SQLDIA
  Set SQLDOC = NewQuery
  Set SQLDIA = NewQuery

  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add("SELECT SOLICITAUX FROM SAM_SOLICITAUX_BENEFICIO WHERE HANDLE = :HSOLICITAUXBEN")
  SQLDIA.ParamByName("HSOLICITAUXBEN").AsInteger = CurrentQuery.FieldByName("SOLICITAUXBENEFICIO").AsInteger
  SQLDIA.Active = True

  SQLDOC.Active = False
  SQLDOC.Clear
  SQLDOC.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
  SQLDOC.ParamByName("HSOLICITAUX").AsInteger = SQLDIA.FieldByName("SOLICITAUX").AsInteger
  SQLDOC.Active = True

  If SQLDOC.FieldByName("SITUACAO").AsString <> "A" Then
    DATAENTREGA.ReadOnly = True
    TIPODOCAUXBENEFICIO.ReadOnly = True
  Else
    DATAENTREGA.ReadOnly = False
    TIPODOCAUXBENEFICIO.ReadOnly = False
  End If
End Sub

Public Sub TABLE_AfterScroll()

  Exit Sub

  Dim SQLDOC As Object
  Dim SQLDIA As Object
  Dim SQLVER As Object
  Set SQLDOC = NewQuery
  Set SQLDIA = NewQuery
  Set SQLVER = NewQuery

  SQLDIA.Active = False
  SQLDIA.Clear
  SQLDIA.Add("SELECT SOLICITAUX FROM SAM_SOLICITAUX_BENEFICIO WHERE HANDLE = :HSOLICITAUXBEN")
  SQLDIA.ParamByName("HSOLICITAUXBEN").AsInteger = CurrentQuery.FieldByName("SOLICITAUXBENEFICIO").AsInteger
  SQLDIA.Active = True

  SQLDOC.Active = False
  SQLDOC.Clear
  SQLDOC.Add("SELECT SITUACAO FROM SAM_SOLICITAUX WHERE HANDLE = :HSOLICITAUX")
  SQLDOC.ParamByName("HSOLICITAUX").AsInteger = SQLDIA.FieldByName("SOLICITAUX").AsInteger
  SQLDOC.Active = True

  If SQLDOC.FieldByName("SITUACAO").AsString <> "A" Then
    DATAENTREGA.ReadOnly = True
    TIPODOCAUXBENEFICIO.ReadOnly = True
  Else
    DATAENTREGA.ReadOnly = False
    TIPODOCAUXBENEFICIO.ReadOnly = False
  End If

  SQLVER.Active = False
  SQLVER.Clear
  SQLVER.Add(" SELECT COUNT(A.HANDLE) AS TOTAL")
  SQLVER.Add(" FROM SAM_SOLICITAUX_BENEFICIO A, ")
  SQLVER.Add("      SAM_SOLICITAUX_BENEFICIO_DIA B")
  SQLVER.Add(" WHERE (VALORPRESTCONTAS IS NOT NULL Or VALORPRESTCONTAS <> 0)")
  SQLVER.Add(" AND A.HANDLE = :HSOLICITAUX")
  SQLVER.Add(" AND B.SOLICITAUXBENEFICIO = A.HANDLE")
  SQLVER.ParamByName("HSOLICITAUX").AsInteger = CurrentQuery.FieldByName("SOLICITAUXBENEFICIO").AsInteger
  SQLVER.Active = True

  If SQLVER.FieldByName("TOTAL").AsInteger <> 0 Then
    MsgBox "Não é possível alterar, prestação de contas efetuada !"
    RefreshNodesWithTable("SAM_SOLICITAUX_BENEFICIO_DIA")
    Exit Sub
  End If

  SQLDOC.Active = False
  SQLDIA.Active = False
  SQLVER.Active = False
  Set SQLDOC = Nothing
  Set SQLDIA = Nothing
  Set SQLVER = Nothing
End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser

End Sub

