'HASH: AC1A47DF5D2C78EE1FB04908E59D7E67

Public Sub RELATORIO_OnExit()
  If (CurrentQuery.FieldByName("TITULO").AsString = "") Then
    CurrentQuery.FieldByName("TITULO").AsString = RELATORIO.Text
  End If
End Sub

