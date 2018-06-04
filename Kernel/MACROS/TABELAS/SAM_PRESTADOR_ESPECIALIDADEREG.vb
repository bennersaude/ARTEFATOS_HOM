'HASH: 78E9917AF9EFCE46C2AABBE660D12C0A
'#Uses "*liberaEspecialidade"
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If liberaEspecialidade <> "" Then
		PRESTADORESPECIALIDADEGRP.ReadOnly = True
	Else
		PRESTADORESPECIALIDADEGRP.ReadOnly = False
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg<>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg<>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Msg = LiberaEspecialidade

	If Msg<>"" Then
		CanContinue = False
		bsShowMessage(Msg, "E")
	End If
End Sub
