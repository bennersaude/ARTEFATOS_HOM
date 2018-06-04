'HASH: 5729D80E79A562FB1F659E51508013CB
 
Public Sub BOTAOIMPRIMIRIR_OnClick()
	Dim qRelatorio As BPesquisa
	Set qRelatorio = NewQuery

	SessionVar("DMEDCPF") = CurrentQuery.FieldByName("CPF").AsString
	SessionVar("DMEDANO") = Str(RecordHandleOfTable("SFN_DMEDANOCALENDARIO"))

	With qRelatorio
		.Active = False
		.Add(" SELECT RELATORIOIR      ")
		.Add("   FROM SFN_PARAMETROSIR ")
		.Active = True
	End With

	If qRelatorio.FieldByName("RELATORIOIR").AsInteger > 0 Then
		ReportPreview(qRelatorio.FieldByName("RELATORIOIR").AsInteger, "",False , False)
	Else
		MsgBox("Relatório de IR não parametrizado.")
	End If

	Set qRelatorio = Nothing

End Sub

Public Sub TABLE_AfterScroll()
	BOTAOIMPRIMIRIR.Visible = VisibleMode
End Sub
