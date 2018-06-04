'HASH: DAAC82FB25802C4FF0C834ECB6098EB2
Option Explicit


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("LER").AsString = "N") Then
    CurrentQuery.FieldByName("ALTERAR").AsString = "N"
    CurrentQuery.FieldByName("EXCLUIR").AsString = "N"
    CurrentQuery.FieldByName("INCLUIR").AsString = "N"
  End If
End Sub
