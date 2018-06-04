'HASH: 26A549251A15E249F4EC295956742EA1
'Macro: SFN_CONTABHIST


Public Sub BOTAOCAMPOS_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("Financeiro.Campos")
  interface.Exec(CurrentSystem, True, 0)
  Set interface = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOCAMPOS" Then
		BOTAOCAMPOS_OnClick
	End If
End Sub
