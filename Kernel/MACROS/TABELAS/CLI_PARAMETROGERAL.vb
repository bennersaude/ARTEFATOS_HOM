'HASH: 4BBB65337D5478E6E72C5ACD38CCD082
'Macro: CLI_PARAMETROGERAL
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

'SMS 46142
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If Not ((CurrentQuery.FieldByName("BENEFICIARIO").AsBoolean) Or (CurrentQuery.FieldByName("NAOBENEFICIARIO").AsBoolean) Or (CurrentQuery.FieldByName("BENEFICIARIONOVO").AsBoolean)) Then
    bsShowMessage("Pelo menos uma das opções do grupo 'Permitir agendamento para' deve estar marcado", "E")
    CanContinue = False
  End If
'FIM SMS 46142
End Sub
