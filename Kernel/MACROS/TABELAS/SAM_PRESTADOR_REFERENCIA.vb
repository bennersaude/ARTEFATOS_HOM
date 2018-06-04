'HASH: 758923BBA6F2146760E1299EE5983768
'#Uses "*bsShowMessage"

Public Sub BOTAORENOVAR_OnClick()
  Dim vDataStr As String
  Dim vData As Date
  Dim q1 As Object

  vDataStr = InputBox("Renovação de referenciamento", "Informar nova data final", CurrentQuery.FieldByName("DATAFINAL").AsString)

  If vDataStr <> "" Then

    If CDate(vDataStr) < CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
        bsShowMessage("A nova data final deve ser maior que a data final atual !","E")
      Exit Sub
    End If

    Set q1 = NewQuery
    If Not InTransaction Then StartTransaction
    q1.Add("INSERT INTO SAM_PRESTADOR_REFERENCIA_RENOV                                            ")
    q1.Add("   (HANDLE,PRESTADORREFERENCIA,DATARENOVACAO,USUARIORENOVACAO,DATAFINALANTERIOR)      ")
    q1.Add("VALUES                                                                                ")
    q1.Add("   (:HANDLE,:PRESTADORREFERENCIA,:DATARENOVACAO,:USUARIORENOVACAO,:DATAFINALANTERIOR) ")
    q1.ParamByName("HANDLE").Value = NewHandle("SAM_PRESTADOR_REFERENCIA_RENOV")
    q1.ParamByName("PRESTADORREFERENCIA").Value = CurrentQuery.FieldByName("HANDLE").Value
    q1.ParamByName("DATARENOVACAO").Value = ServerNow
    q1.ParamByName("USUARIORENOVACAO").Value = CurrentUser
    q1.ParamByName("DATAFINALANTERIOR").Value = CurrentQuery.FieldByName("DATAFINAL").Value
    q1.ExecSQL
    If InTransaction Then Commit

    CurrentQuery.Edit
    CurrentQuery.FieldByName("DATAFINAL").Value = CDate(vDataStr)
    CurrentQuery.Post
  End If

End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  DATAINICIAL.ReadOnly = True
  DATAFINAL.ReadOnly = True
  MOTIVOREFERENCIAMENTO.ReadOnly = True
  OBSERVACAO.ReadOnly = True

  If Not CurrentQuery.FieldByName("OBSERVACAO").IsNull Then
    OBSERVACAO.ReadOnly = True
  Else
    OBSERVACAO.ReadOnly = False
  End If
End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If (CommandID = BOTAORENOVAR) Then
		BOTAORENOVAR_OnClick
	End If
End Sub
