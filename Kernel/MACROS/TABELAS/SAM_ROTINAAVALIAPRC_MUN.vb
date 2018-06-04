'HASH: 64D703F10B32F1B74CA48FA5F589934C
'#Uses "*bsShowMessage"

Public Sub ExcluirdePrestadorAtePrestador()

  If Not InTransaction Then
    StartTransaction
  End If


  Dim qd As Object

  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_PRE_VAL                      ")
  qd.Add("      WHERE HANDLE IN (SELECT A.HANDLE                       ")
  qd.Add("                         FROM SAM_ROTINAAVALIAPRC_PRE_VAL A ,")
  qd.Add("                              SAM_ROTINAAVALIAPRC_PRE B,     ")
  qd.Add("                              SAM_ROTINAAVALIAPRC_MUN C,     ")
  qd.Add("                              SAM_ROTINAAVALIAPRC_EST D,     ")
  qd.Add("                              SAM_ROTINAAVALIAPRC E          ")
  qd.Add("                        WHERE E.HANDLE = :HANDLE             ")
  qd.Add("                          And D.AVALIAPRC = E.HANDLE         ")
  qd.Add("                          And C.AVALIAPRC = D.HANDLE         ")
  qd.Add("                          And B.AVALIAPRC = C.HANDLE         ")
  qd.Add("                          And A.AVALIAPRC = B.HANDLE)        ")
  qd.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_ROTINAAVALIAPRC")
  qd.ExecSQL

  Set qd = Nothing
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_PRE                      ")
  qd.Add("      WHERE HANDLE IN (SELECT A.HANDLE                   ")
  qd.Add("                         FROM SAM_ROTINAAVALIAPRC_PRE A ,")
  qd.Add("                              SAM_ROTINAAVALIAPRC_MUN B, ")
  qd.Add("                              SAM_ROTINAAVALIAPRC_EST C, ")
  qd.Add("                              SAM_ROTINAAVALIAPRC D      ")
  qd.Add("                        WHERE D.HANDLE = :HANDLE         ")
  qd.Add("                          And C.AVALIAPRC = D.HANDLE     ")
  qd.Add("                          And B.AVALIAPRC = C.HANDLE     ")
  qd.Add("                          And A.AVALIAPRC = B.HANDLE)    ")
  qd.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_ROTINAAVALIAPRC")
  qd.ExecSQL

  Set qd = Nothing

  Dim qu As Object
  Set qu = NewQuery
  qu.Add("UPDATE SAM_ROTINAAVALIAPRC SET SITUACAO = 'A' WHERE HANDLE = :HANDLE")
  qu.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qu.ExecSQL
  Set qu = Nothing

  If InTransaction Then
    Commit
  End If

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_MUN")

End Sub

Public Sub BOTAOCANCELAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").Value = "A" Then
    bsShowMessage("Exclusão não efetuada. Avaliação Aberta.", "I")
    Exit Sub
  End If

  Dim qm As Object

  Set qm = NewQuery

  qm.Add("SELECT * FROM SAM_ROTINAAVALIAPRC_MUN_VAL WHERE AVALIAPRC=:AVALIAPRC")
  qm.ParamByName("AVALIAPRC").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qm.Active = True

  If qm.EOF Then
    bsShowMessage("Não há Avaliação para este Municípios.", "I")
    Exit Sub
  End If
  If bsShowMessage("Excluir Avaliação do Município?", "Q") = vbYes Then

    If Not InTransaction Then StartTransaction

    Dim q1 As Object
    Set q1 = NewQuery

    q1.Add("DELETE FROM SAM_ROTINAAVALIAPRC_MUN_VAL WHERE AVALIAPRC = :AVALIAPRC")
    q1.ParamByName("AVALIAPRC").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    q1.ExecSQL

    Dim q2 As Object

    Set q2 = NewQuery

    q2.Add("UPDATE SAM_ROTINAAVALIAPRC_MUN SET SITUACAO = 'A' WHERE HANDLE = :HANDLEAVALIA")
    q2.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    q2.ExecSQL

    If InTransaction Then Commit

    RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_MUN")

    Set q2 = Nothing

  End If

  Set q1 = Nothing

End Sub

Public Sub BOTAOCANCELARPRESTADOR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If

  Dim qp As Object

  Set qp = NewQuery

  qp.Add("SELECT * FROM SAM_ROTINAAVALIAPRC_PRE WHERE AVALIAPRC=:HANDLE")
  qp.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qp.Active = True

  If qp.EOF Then
    bsShowMessage("Não há Avaliação para Prestadores neste Municípios.", "I")
    Exit Sub
  End If

  If bsShowMessage("Excluir Prestadores e suas Avaliações?", "Q") = vbYes Then

    ExcluirdePrestadorAtePrestador

    RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_PRE")

  End If


End Sub

Public Sub BOTAOGERAL_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    bsShowMessage("Avaliação já Processada.", "I")
    Exit Sub
  End If



  'CHAMA DLL PASSANDO OS PARAMENTROS  -DLLROTAVALIAPRC.Geral

  Dim vmsgRetorno As String
  Set Obj = CreateBennerObject("BSINTERFACE0010.RotinaGerarAvaliacao")
  Obj.AvaliaPrecoGerarAvaliacao(CurrentSystem, RecordHandleOfTable("SAM_ROTINAAVALIAPRC"), "M", CurrentQuery.FieldByName("AVALIAPRC").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, CurrentQuery.FieldByName("ESTADOS").AsInteger, CurrentQuery.FieldByName("MUNICIPIOS").AsInteger, 0)' pedir confirmação

  If vmsgRetorno <> "" Then
    bsShowMessage(vmsgRetorno, "E")
  End If

 ' Set Obj = CreateBennerObject("SAMROTAVALIAPRC.geral")
 ' Obj.executar(CurrentSystem, RecordHandleOfTable("SAM_ROTINAAVALIAPRC"), "M", CurrentQuery.FieldByName("AVALIAPRC").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, CurrentQuery.FieldByName("ESTADOS").AsInteger, CurrentQuery.FieldByName("MUNICIPIOS").AsInteger, 0)' pedir confirmação


  CurrentQuery.Active = False
  CurrentQuery.Active = True

  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_MUN")

End Sub

Public Sub BOTAOPRESTADOR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If

  Dim q2 As Object

  Set q2 = NewQuery

  q2.Add("SELECT * FROM SAM_ROTINAAVALIAPRC_PRE WHERE AVALIAPRC = :HANDLEAVALIA")
  q2.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  q2.Active = True

  If Not q2.EOF Then
    bsShowMessage("Avaliação de Prestadores já Processada.", "I")
    Exit Sub
  End If

  Set q2 = Nothing

  'CHAMA DLL PASSANDO OS PARAMENTROS  -DLLROTAVALIAPRC.Geral
  '            Principal(pHRotAvaliaPrc:Integer;

  'Set Obj = CreateBennerObject("SAMROTAVALIAPRC.geral")
  'Obj.executar(CurrentSystem, RecordHandleOfTable("SAM_ROTINAAVALIAPRC"), "P", CurrentQuery.FieldByName("AVALIAPRC").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, CurrentQuery.FieldByName("ESTADOS").AsInteger, CurrentQuery.FieldByName("MUNICIPIOS").AsInteger, 0)' pedir confirmação

  Dim vmsgRetorno As String
  Set Obj = CreateBennerObject("BSINTERFACE0010.RotinaGerarAvaliacao")
  Obj.AvaliaPrecoGerarPrestador(CurrentSystem, RecordHandleOfTable("SAM_ROTINAAVALIAPRC"), "P", CurrentQuery.FieldByName("AVALIAPRC").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, CurrentQuery.FieldByName("ESTADOS").AsInteger, CurrentQuery.FieldByName("MUNICIPIOS").AsInteger, 0)' pedir confirmação

  If vmsgRetorno <> "" Then
    bsShowMessage(vmsgRetorno, "E")
  End If


  CurrentQuery.Active = False
  CurrentQuery.Active = True

  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_PRE")

End Sub

Public Sub TABLE_AfterPost()

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    MUNICIPIOS.ReadOnly = True
  Else
    MUNICIPIOS.ReadOnly = False
  End If

  Set qp = NewQuery

  qp.Add("SELECT * FROM SAM_ROTINAAVALIAPRC_PRE WHERE AVALIAPRC=:HANDLE")
  qp.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qp.Active = True

  If Not qp.EOF Then
    MUNICIPIOS.ReadOnly = True
  End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  'If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
  '  CanContinue =False
  '  MsgBox "Operação cancelada. Usuário diferente"
  '  Exit Sub
  'End If

  bsShowMessage("Para excluir, deve-se usar o botão cancelar no processo acima.", "E")
  CanContinue = False

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    CanContinue = False
    bsShowMessage("Registro finalizado não pode ser alterado!", "E")
    Exit Sub
  End If
  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário diferente", "E")
    Exit Sub
  End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    CanContinue = False
    bsShowMessage("Alteração não Efetuada. Avaliação com situação Processada.", "E")
    Exit Sub
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOCANCELARPRESTADOR"
			BOTAOCANCELARPRESTADOR_OnClick
		Case "BOTAOGERAL"
			BOTAOGERAL_OnClick
		Case "BOTAOPRESTADOR"
			BOTAOPRESTADOR_OnClick
	End Select
End Sub
