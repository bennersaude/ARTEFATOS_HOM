'HASH: 9FAF34E20A47D7424538E9FC71A25708
'#uses "*bsShowMessage"

Public Sub CONCEDERATODOS_OnClick()
  On Error GoTo Erro
    If CurrentQuery.State <> 1 Then Exit Sub
    Dim Q As Object
    Dim Q2 As Object
    Set Q = NewQuery
    Set Q2 = NewQuery
    Q.Add("SELECT HANDLE FROM Z_GRUPOUSUARIOS WHERE HANDLE NOT IN " + _
          "(SELECT USUARIO FROM Z_GRUPOUSUARIOALCADAS WHERE ALCADA = " + _
          CurrentQuery.FieldByName("HANDLE").AsString + ")")
    Q.Active = True
    While Not Q.EOF
      Q2.Clear

      If Not InTransaction Then StartTransaction
	    Q2.Add("INSERT INTO Z_GRUPOUSUARIOALCADAS (HANDLE, USUARIO, ALCADA) " + _
           	" VALUES (" + CStr(NewHandle("Z_GRUPOUSUARIOALCADAS")) + ", " + _
           	Q.FieldByName("HANDLE").AsString + ", " + _
           	CurrentQuery.FieldByName("HANDLE").AsString + ")")
    	Q2.ExecSQL
      If InTransaction Then Commit
      Q.Next
    Wend
    Q.Active = False
    Set Q2 = Nothing
    Set Q = Nothing

    bsShowMessage("Processo concluído com sucesso!", "I")

    Exit Sub

  Erro:
    If InTransaction Then
      Rollback
    End If
    bsShowMessage("Erro no processamento: " + Err.Description, "I")
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "CONCEDERATODOS" Then
		CONCEDERATODOS_OnClick
	End If
End Sub
