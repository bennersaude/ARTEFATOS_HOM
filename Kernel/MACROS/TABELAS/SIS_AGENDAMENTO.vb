'HASH: 65D6FB38051CC8E3C83A6F630A5992B8
'#Uses "*bsShowMessage"

Public Sub BOTAOAGENDAR_OnClick()
  Dim qr As Object
  Set qr = NewQuery
  If CurrentQuery.FieldByName("SITUACAO").AsString = "D" Then
    bsShowMessage("Tarefa já está agendada !", "I")
  Else
    qr.Clear
	If Not InTransaction Then StartTransaction
     qr.Add("UPDATE SIS_AGENDAMENTO SET SITUACAO = 'D', USUARIOCANCELAMENTO = NULL, DATACANCELAMENTO = NULL  WHERE HANDLE = :pHANDLE")
     qr.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
     qr.ExecSQL
     Set qr = Nothing
  	If InTransaction Then Commit
    SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Set qr = NewQuery
  If CurrentQuery.FieldByName("SITUACAO").AsString = "C" Then
    bsShowMessage("Tarefa já está CANCELADA !", "E")
  Else
    qr.Clear
     If Not InTransaction Then StartTransaction
    	qr.Add("UPDATE SIS_AGENDAMENTO SET SITUACAO = 'C', USUARIOCANCELAMENTO = :pUSUARIOCANCELAMENTO, DATACANCELAMENTO = :pDATACANCELAMENTO WHERE HANDLE = :pHANDLE")
    	qr.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    	qr.ParamByName("pUSUARIOCANCELAMENTO").AsInteger = CurrentUser
    	qr.ParamByName("pDATACANCELAMENTO").AsDateTime = ServerNow
    	qr.ExecSQL
    	Set qr = Nothing
	  If InTransaction Then Commit
    If VisibleMode Then
      SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
    End If
  End If
End Sub

Public Sub TABLE_AfterCommitted()
  Dim Obj As Object
  Set Obj = CreateBennerObject("BsProcesso.Schedule")
  Obj.Executar(CurrentSystem, CurrentQuery.FieldByName("PROCESSO").AsInteger)
  Set Obj = Nothing
End Sub

Public Sub TABLE_AfterPost()

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Obj As Object
  Set Obj = CreateBennerObject("BsProcesso.Schedule")
  Obj.Deletar(CurrentSystem, CurrentQuery.FieldByName("SCHEDULE").AsInteger)
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOAGENDAR"
			BOTAOAGENDAR_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
	End Select
End Sub
