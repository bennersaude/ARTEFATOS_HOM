'HASH: 7B8832C2C7CE88D4C83B035C4843C055
 
'#Uses "*LimpaEspaco"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not CurrentQuery.FieldByName("NOTAFISCAL").IsNull Then
    CurrentQuery.FieldByName("NOTAFISCAL").AsString = LimpaEspaco(CurrentQuery.FieldByName("NOTAFISCAL").AsString)
  End If
End Sub
