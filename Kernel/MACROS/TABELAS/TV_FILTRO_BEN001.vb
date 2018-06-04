'HASH: D3ED4763BC8ED222A4ECDE58FCF02DA1
Public Sub TABLE_OnVirtualValue(ByVal MacroName As String, ResultValue As String)

  If (IsNumeric(SessionVar("ROTINASIMULACAOBEN001"))) Then
    CurrentQuery.FieldByName("ROTINA").AsInteger = CLng(SessionVar("ROTINASIMULACAOBEN001"))
  End If

End Sub
