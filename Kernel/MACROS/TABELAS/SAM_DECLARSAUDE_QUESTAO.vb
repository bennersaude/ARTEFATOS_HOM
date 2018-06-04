'HASH: F152E5CCF4B3F8D6D37F6FE356FC320A
'MACRO SAM_DECLARSAUDE_QUESTAO - Durval

Option Explicit
Dim SQL As Object

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT * FROM SAM_DECLARSAUDE_QUESTAO ")
  SQL.Add("WHERE DECLARSAUDE = :DECLARACAO ")
  SQL.Add("AND CODIGO = :CODIGO ")
  SQL.ParamByName("DECLARACAO").AsInteger = CurrentQuery.FieldByName("DECLARSAUDE").AsInteger
  SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.Active = True

  If CurrentQuery.State = 3 Then
    If Not SQL.FieldByName("CODIGO").IsNull Then
      MsgBox ("Este código já existe nesta Declaração!")
      CODIGO.SetFocus
      CanContinue = False
    End If
  End If

  Set SQL = Nothing
End Sub

