'HASH: EFD22BFA41E1E78F6E4C393CD20D1394
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

  '#Uses "*ProcuraEvento"
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text) ' só último nível
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery
	SQL.Add("SELECT COUNT(1)  NUM         ")
	SQL.Add("  FROM SAM_TIPOAUTORIZ_EVENTO")
	SQL.Add(" WHERE HANDLE <> :HANDLE")
	SQL.Add("   AND EVENTO = :EVENTO")
	SQL.Add("   AND TIPOAUTORIZ = :TIPOAUTORIZ")
	SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
	SQL.ParamByName("TIPOAUTORIZ").AsInteger = CurrentQuery.FieldByName("TIPOAUTORIZ").AsInteger
	SQL.Active = True

	If SQL.FieldByName("NUM").AsInteger > 0 Then
	    bsShowMessage("Evento já inserido.", "E")
		CanContinue = False
	End If

	Set SQL = Nothing
End Sub
