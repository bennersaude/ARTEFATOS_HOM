'HASH: 0889EA6F15B7398C041AF6B14F313737
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub EVENTOPREDECESSOR_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False

  vHandle = ProcuraEvento(True, EVENTOPREDECESSOR.Text)

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOPREDECESSOR").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterScroll()

	If (VisibleMode Or WebMode) Then
  		EVENTOPREDECESSOR.WebLocalWhere = "A.ULTIMONIVEL = 'S' AND A.HANDLE <> " + Str(RecordHandleOfTable("SAM_TGE"))
  	Else
  		EVENTOPREDECESSOR.WebLocalWhere = "A.ULTIMONIVEL = 'S' AND A.HANDLE <> " + CurrentQuery.FieldByName("EVENTO").AsString
  	End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim handleEvento As Long

	If (VisibleMode Or WebMode) Then
		handleEvento = RecordHandleOfTable("SAM_TGE")
	Else
		handleEvento = CurrentQuery.FieldByName("EVENTO").AsInteger
	End If

	If (CurrentQuery.FieldByName("EVENTOPREDECESSOR").AsInteger = handleEvento) Then
		CanContinue = False
		bsShowMessage("Evento predecessor deve ser diferente do próprio evento!", "E")
  	End If

End Sub
