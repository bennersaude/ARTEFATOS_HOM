'HASH: 7C63FDE6F9AF5F3D082E267176B40C23
'#Uses "*bsShowMessage"

Public Sub BOTAOCOPIARROTINA_AfterOnClick()
  RefreshNodesWithTable("SAM_ROTNOVASADESOES")
End Sub

Public Sub BOTAOIMPRIMIR_BeforeOnClick(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("SITUACAO").AsString <> "5") Then
		bsShowMessage("Rotina ainda não processada", "E")
		CanContinue = False
	End If
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'BEN001'")
  SQL.Active = True
  SessionVar("ROTINASIMULACAOBEN001") = CurrentQuery.FieldByName("HANDLE").AsString

  ReportPreview(SQL.FieldByName("HANDLE").Value, "", True, False)

  SessionVar("ROTINASIMULACAOBEN001") = ""

  Set SQL = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  RecordReadOnly = CurrentQuery.FieldByName("SITUACAO").AsString = "5"
End Sub
