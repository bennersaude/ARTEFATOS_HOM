'HASH: 109B634E12CA6CA88BD42E5E033A1143
Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  Dim Apaga As Object
  Set Apaga = NewQuery
  Dim VHandleODonto As Long

  If Not InTransaction Then StartTransaction

  VHandleODonto = CurrentQuery.FieldByName("ODONTOGRAMA").AsInteger

  Apaga.Clear
  Apaga.Add("DELETE FROM SAM_GUIA_ODONTOGRAMA")
  Apaga.Add(" WHERE ODONTOGRAMA = :ODONTOGRAMA")
  Apaga.ParamByName("ODONTOGRAMA").AsInteger = VHandleODonto
  Apaga.ExecSQL

  Apaga.Clear
  Apaga.Add("DELETE FROM CLI_ODONTOGRAMA")
  Apaga.Add(" WHERE HANDLE = :HANDLE")
  Apaga.ParamByName("HANDLE").AsInteger = VHandleODonto
  Apaga.ExecSQL
  If InTransaction Then Commit

  RefreshNodesWithTable("SAM_GUIA_ODONTOGRAMA")

  Set Apaga = Nothing
End Sub

Public Sub BOTAOCONFIRMAR_OnClick()
  Dim SQL As Object
  Dim Apaga As Object
  Set SQL = NewQuery
  Set Apaga = NewQuery
  Dim VHandleODonto As Long

  If Not InTransaction Then StartTransaction

  VHandleODonto = CurrentQuery.FieldByName("ODONTOGRAMA").AsInteger

  Apaga.Clear
  Apaga.Add("DELETE FROM SAM_GUIA_ODONTOGRAMA")
  Apaga.Add(" WHERE ODONTOGRAMA = :ODONTOGRAMA")
  Apaga.ParamByName("ODONTOGRAMA").AsInteger = VHandleODonto
  Apaga.ExecSQL

  SQL.Clear
  SQL.Add("UPDATE CLI_ODONTOGRAMA")
  SQL.Add("   SET HISTORICO = :HISTORICO,")
  SQL.Add("       SITUACAO  = :SITUACAO")
  SQL.Add(" WHERE HANDLE    = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ODONTOGRAMA").AsInteger
  SQL.ParamByName("HISTORICO").AsString = "S"
  SQL.ParamByName("SITUACAO").AsString = "RO"
  SQL.ExecSQL

  If InTransaction Then Commit

  RefreshNodesWithTable("SAM_GUIA_ODONTOGRAMA")

  Set Apaga = Nothing
  Set SQL = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  If (CurrentQuery.FieldByName("FACE").AsString = "") Or CurrentQuery.FieldByName("FACE").IsNull Then
    FACE.Visible = False
  Else
    FACE.Visible = True
  End If
End Sub

