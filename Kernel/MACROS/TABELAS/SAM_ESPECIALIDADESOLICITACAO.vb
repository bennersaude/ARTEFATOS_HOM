'HASH: A73B9EAE7B0E511AAB571FEC4C846B22
'Macro: SAM_ESPECIALIDADESOLICITACAO
'#Uses "*ProcuraEvento"
'#uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If

		EVENTO.WebLocalWhere = "A.INATIVO = 'N' " 'Luciano T. Alberti - SMS 95290 - 01/04/2008
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		If WebVisionCode = "V_SAM_ESPECIALIDADESOLICITAC_592" Then
			ESPECIALIDADE.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		If WebVisionCode = "V_SAM_ESPECIALIDADESOLICITAC_592" Then
			ESPECIALIDADE.ReadOnly = True
		End If
	End If
End Sub
