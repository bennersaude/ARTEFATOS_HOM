'HASH: C9ADE90FB29771737D6D7A526B4C1DB4
'#Uses "*bsShowMessage"
'Sis_StoreProcedure

Public Sub BOTAOVERIFICA_OnClick()
  Dim Obj As Object
  Dim BANCO As String

  BANCO = CurrentQuery.FieldByName("BANCO").AsString

  If BANCO = "O" Then
    BANCO = "ORACLE"
  End If

  BANCO = CurrentQuery.FieldByName("BANCO").AsString

  If BANCO = "C" Then
    BANCO = "CACHE"
  End If

  If BANCO = "M" Then
    BANCO = "MSSQL"
  End If

  If BANCO = "D" Then
    BANCO = "DB2"
  End If

  If InStr(SQLServer, BANCO)Then
    Set Obj = CreateBennerObject("SamUtil.Rotinas")
    Obj.VerificaSP(CurrentSystem)
    Set Obj = Nothing
  Else
    bsShowMessage("O banco em uso é " + SQLServer, "")
  End If


End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOVERIFICA" Then
		BOTAOVERIFICA_OnClick
	End If
End Sub
