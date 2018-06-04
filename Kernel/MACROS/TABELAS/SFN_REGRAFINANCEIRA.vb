'HASH: D950D90C5002ED0A13F49A71EEA4BFEA
'Macro: SFN_REGRAFINANCEIRA
'#Uses "*bsShowMessage"

Public Sub CALCULAR_OnClick()


 If CurrentQuery.State <> 1 Then
    bsShowMessage("A regra financeira não pode estar em edição.", "I")
    Exit Sub
  End If

  If VisibleMode Then

	  Dim INTERFACE0002 As Object
	  Dim vsMensagem As String
	  Dim vcContainer As CSDContainer

	  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")


	  INTERFACE0002.Exec(CurrentSystem, _
						   1, _
						   "TV_FORM0110", _
						   "Consulta financeira",  _
						   0, _
						   450, _
						   350, _
						   False, _
						   vsMensagem, _
						   vcContainer)
	  Set INTERFACE0002 = Nothing
	  Set vcContainer = Nothing

  End If

End Sub

Public Sub TABLE_AfterScroll()
  UserVar("HANDLE_REGRAFINANCEIRA") = CurrentQuery.FieldByName("HANDLE").Value
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = "CALCULAR") Then
		CALCULAR_OnClick
	End If
End Sub
