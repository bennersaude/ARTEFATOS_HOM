'HASH: E330B362E21F755112B7354E0B5E654F
Option Explicit





Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("LER").AsString = "N") Then
    CurrentQuery.FieldByName("ALTERAR").AsString = "N"
    CurrentQuery.FieldByName("EXCLUIR").AsString = "N"
    CurrentQuery.FieldByName("INCLUIR").AsString = "N"
  End If
End Sub
