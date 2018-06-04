'HASH: 606ABFE26F7C843E596E1C71575EB3DC
Option Explicit





Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (CurrentQuery.FieldByName("LER").AsString = "N") Then
    CurrentQuery.FieldByName("ALTERAR").AsString = "N"
  End If
End Sub
