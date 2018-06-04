'HASH: 6DF8DD70C384F55AA9FD3FB916D07D4D
Option Explicit

Public Sub TABLE_AfterScroll()

  TIPODETAXA.ReadOnly = False
  DATAINICIAL.ReadOnly = False

  If ((CurrentQuery.State <> 3) And (CurrentQuery.FieldByName("JAUTILIZADA").AsString = "S")) Then
      TIPODETAXA.ReadOnly = True
      DATAINICIAL.ReadOnly = True
  End If

End Sub
