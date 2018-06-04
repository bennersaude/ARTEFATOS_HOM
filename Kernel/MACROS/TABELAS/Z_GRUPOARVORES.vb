'HASH: C47F240C1EFBED98CEEF195A39A4E934
Option Explicit




Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("LER").AsString = "N") Then
    CurrentQuery.FieldByName("ALTERAR").AsString = "N"
    CurrentQuery.FieldByName("EXCLUIR").AsString = "N"
    CurrentQuery.FieldByName("INCLUIR").AsString = "N"
  End If
End Sub
