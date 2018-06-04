'HASH: 455A9BB5FAA0024AB52D5943B64FA3B9
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABELABASE_OnChange()
	If (CurrentQuery.FieldByName("TABELABASE").IsNull And CurrentQuery.State <> 1 ) Then
		PERCENTUALNEGOCIADO.ReadOnly = True
		CurrentQuery.FieldByName("PERCENTUALNEGOCIADO").Value = Null
	Else
		PERCENTUALNEGOCIADO.ReadOnly = False
	End If
End Sub

Public Sub TABLE_AfterScroll()
	TABELABASE_OnChange
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem,"A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem,"I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
