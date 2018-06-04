'HASH: 11ECBB6F02F9DD9F370653F5EADC66FB
Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("PERCENTUAL").AsFloat > 100 Then
    MsgBox "O Percentual não pode ser maior que 100 %"
    CanContinue = False
  End If
End Sub

