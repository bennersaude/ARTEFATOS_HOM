'HASH: 13C3D430E4EC63BBF53BA74EE6393FFD

'MACRO SAM_DECLARSAUDE - Durval
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT CODIGO FROM SAM_DECLARSAUDE ")
  SQL.Add("WHERE CODIGO = :CODIGO ")
  SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
  SQL.Active = True

  If CurrentQuery.State = 3 Then
    If Not SQL.FieldByName("CODIGO").IsNull Then
      bsShowMessage("Já existe Declaração com este código!", "E")
      CODIGO.SetFocus
      CanContinue = False
    End If
  End If

  Set SQL = Nothing
End Sub

