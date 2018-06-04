'HASH: D64A4D34B5ECCEC577838AC08F348094
Option Explicit

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim sql As BPesquisa
  Set sql = NewQuery
  sql.Add("SELECT COUNT(1) NREC FROM SAM_GRUPOUSUARIO WHERE USUARIO=:U")
  sql.ParamByName("U").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
  sql.Active=True
  If sql.FieldByName("NREC").AsInteger>0 Then
    CanContinue=False
    CancelDescription="Já existe registro para este usuário"
  End If
  Set sql=Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
    Dim QueryNiveis As BPesquisa
    Set QueryNiveis = NewQuery
    QueryNiveis.Clear
    QueryNiveis.Add("SELECT N.NIVEL FROM SAM_NIVELAUTORIZACAO N WHERE N.HANDLE=:NIVEL")
    QueryNiveis.ParamByName("NIVEL").Value = CurrentQuery.FieldByName("NIVEL").AsInteger
    QueryNiveis.Active = False
    QueryNiveis.Active = True

    CurrentQuery.FieldByName("NIVELAUTORIZACAO").AsInteger = QueryNiveis.FieldByName("NIVEL").AsInteger
    Set QueryNiveis = Nothing
  End If


End Sub

