'HASH: 6ABEE92C43D8057C08D7225CD2C5E1EE

Public Sub INSERIRNIVEL_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("CS.InsertLevel")
  Obj.Exec
  Set Obj = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  If (CurrentQuery.State <> 3) And (CurrentQuery.FieldByName("HANDLE").AsInteger > 0) Then
    Set SQL = NewQuery
    SQL.Clear
    SQL.Active = False
    SQL.Add("SELECT EMPRESA FROM Z_MASCARAS WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
    SQL.Active = True

    If SQL.FieldByName("EMPRESA").IsNull Then
      		SQL.Clear
      		SQL.Add("UPDATE Z_MASCARAS SET EMPRESA=:EMPRESA WHERE HANDLE=:HANDLE")
      		SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      		SQL.ParamByName("EMPRESA").AsInteger = CurrentCompany
      		SQL.ExecSQL
      RefreshNodesWithTable("Z_MASCARAS")
    End If

    Set SQL = Nothing
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "INSERIRNIVEL" Then
		INSERIRNIVEL_OnClick
	End If
End Sub
