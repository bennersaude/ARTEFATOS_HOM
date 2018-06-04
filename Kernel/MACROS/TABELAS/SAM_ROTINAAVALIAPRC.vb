'HASH: 4CDC8B99B05723842A554828727CBCE1
'Macro: SAM_ROTINAAVALIAPRC

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Dim pgESTRUTURAINICIAL As String
Dim pgESTRUTURAFINAL As String

Public Function CheckParametros()

  CheckParametros = False

  pgESTRUTURAINICIAL = ""
  pgESTRUTURAFINAL = ""

  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("Data Inicial Maior que Data Final.", "I")
    Exit Function
  End If

  pgESTRUTURAINICIAL = VERIFICAESTRUTURA(CurrentQuery.FieldByName("EVENTOINICIAL").AsInteger)
  pgESTRUTURAFINAL = VERIFICAESTRUTURA(CurrentQuery.FieldByName("EVENTOFINAL").AsInteger)


  If pgESTRUTURAINICIAL >pgESTRUTURAFINAL Then
    bsShowMessage("Evento Inicial Maior que Evento Final.", "I")
    Exit Function
  End If


  CheckParametros = True

End Function


Public Function VERIFICAESTRUTURA(pEVENTO As Long)As String

  Dim q1 As Object

  Set q1 = NewQuery

  q1.Add("SELECT ESTRUTURA FROM SAM_TGE WHERE HANDLE=:EVENTO")
  q1.ParamByName("EVENTO").Value = pEVENTO
  q1.Active = True

  While Not q1.EOF

    VERIFICAESTRUTURA = q1.FieldByName("ESTRUTURA").AsString
    Set q1 = Nothing
    Exit Function

  Wend

  Set q1 = Nothing

End Function


Public Sub BOTAOCANCELAR_OnClick()
  If bsShowMessage("Excluir Avaliação Geral?", "Q") = vbYes Then
    Dim q1 As Object
    Set q1 = NewQuery

    If Not InTransaction Then StartTransaction

    q1.Add("DELETE from SAM_ROTINAAVALIAPRC_VAL WHERE AVALIAPRC = :HANDLEAVALIA")
    q1.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    q1.ExecSQL
    Set q1 = Nothing

    Dim q2 As Object
    Set q2 = NewQuery
    q2.Add("UPDATE SAM_ROTINAAVALIAPRC SET SITUACAO = 'A' WHERE HANDLE = :HANDLEAVALIA")
    q2.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    q2.ExecSQL
    Set q2 = Nothing

    If InTransaction Then Commit

    RefreshNodesWithTable("SAM_ROTINAAVALIAPRC")

  End If

End Sub

Public Sub BOTAOCANCELARESTADO_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If

  Dim qe As Object

  Set qe = NewQuery

  qe.Add("SELECT * FROM SAM_ROTINAAVALIAPRC_EST WHERE AVALIAPRC=:HANDLE")
  qe.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qe.Active = True


  If Not qe.EOF Then

    If bsShowMessage("Excluir Estados e suas Avaliações.", "Q") = vbYes Then

      ExcluirdeEstadoAtePrestador


      RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_EST")

    End If
  Else
    bsShowMessage("Não há Estados para excluir.", "I")
  End If

  Set qe = Nothing

End Sub

Public Sub ExcluirdeEstadoAtePrestador()
  If Not InTransaction Then
    StartTransaction
  End If

  Dim qd As Object
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_PRE_VAL                   ")
  qd.Add(" WHERE HANDLE IN ( SELECT A.HANDLE                        ")
  qd.Add("                     FROM SAM_ROTINAAVALIAPRC_PRE_VAL A , ")
  qd.Add("                          SAM_ROTINAAVALIAPRC_PRE B,      ")
  qd.Add("                          SAM_ROTINAAVALIAPRC_MUN C,      ")
  qd.Add("                          SAM_ROTINAAVALIAPRC_EST D,      ")
  qd.Add("                          SAM_ROTINAAVALIAPRC E           ")
  qd.Add("                    WHERE E.HANDLE = :HANDLE              ")
  qd.Add("                      And D.AVALIAPRC = E.HANDLE          ")
  qd.Add("                      And C.AVALIAPRC = D.HANDLE          ")
  qd.Add("                      And B.AVALIAPRC = C.HANDLE          ")
  qd.Add("                      And A.AVALIAPRC = B.HANDLE  )       ")
  qd.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qd.ExecSQL

  Set qd = Nothing
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_PRE                   ")
  qd.Add("      WHERE HANDLE IN (SELECT A.HANDLE                   ")
  qd.Add("                         FROM SAM_ROTINAAVALIAPRC_PRE A ,")
  qd.Add("                              SAM_ROTINAAVALIAPRC_MUN B, ")
  qd.Add("                              SAM_ROTINAAVALIAPRC_EST C, ")
  qd.Add("                              SAM_ROTINAAVALIAPRC D      ")
  qd.Add("                        WHERE D.HANDLE = :HANDLE         ")
  qd.Add("                          And C.AVALIAPRC = D.HANDLE     ")
  qd.Add("                          And B.AVALIAPRC = C.HANDLE     ")
  qd.Add("                          And A.AVALIAPRC = B.HANDLE)    ")
  qd.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qd.ExecSQL

  Set qd = Nothing
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_MUN_VAL                       ")
  qd.Add("      WHERE HANDLE IN ( SELECT B.HANDLE                       ")
  qd.Add("                          FROM SAM_ROTINAAVALIAPRC_MUN_VAL B, ")
  qd.Add("                               SAM_ROTINAAVALIAPRC_MUN C,     ")
  qd.Add("                               SAM_ROTINAAVALIAPRC_EST D,     ")
  qd.Add("                               SAM_ROTINAAVALIAPRC E          ")
  qd.Add("                         WHERE E.HANDLE = :HANDLE             ")
  qd.Add("                           And D.AVALIAPRC = E.HANDLE         ")
  qd.Add("                           And C.AVALIAPRC = D.HANDLE         ")
  qd.Add("                           And B.AVALIAPRC = C.HANDLE)        ")
  qd.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qd.ExecSQL

  Set qd = Nothing
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_MUN                       ")
  qd.Add("      WHERE HANDLE IN ( SELECT B.HANDLE                   ")
  qd.Add("                          FROM SAM_ROTINAAVALIAPRC_MUN B, ")
  qd.Add("                               SAM_ROTINAAVALIAPRC_EST C, ")
  qd.Add("                               SAM_ROTINAAVALIAPRC D      ")
  qd.Add("                         WHERE D.HANDLE = :HANDLE         ")
  qd.Add("                           And C.AVALIAPRC = D.HANDLE     ")
  qd.Add("                           And B.AVALIAPRC = C.HANDLE)    ")
  qd.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qd.ExecSQL

  Set qd = Nothing
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_EST_VAL                       ")
  qd.Add("      WHERE HANDLE IN ( SELECT C.HANDLE                       ")
  qd.Add("                          FROM SAM_ROTINAAVALIAPRC_EST_VAL C, ")
  qd.Add("                               SAM_ROTINAAVALIAPRC_EST D,     ")
  qd.Add("                               SAM_ROTINAAVALIAPRC E          ")
  qd.Add("                         WHERE E.HANDLE = :HANDLE             ")
  qd.Add("                           AND D.AVALIAPRC = E.HANDLE         ")
  qd.Add("                           AND C.AVALIAPRC = D.HANDLE)        ")
  qd.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qd.ExecSQL

  Set qd = Nothing
  Set qd = NewQuery

  qd.Add("DELETE FROM SAM_ROTINAAVALIAPRC_EST                       ")
  qd.Add("      WHERE HANDLE IN ( SELECT C.HANDLE                   ")
  qd.Add("                          FROM SAM_ROTINAAVALIAPRC_EST C, ")
  qd.Add("                               SAM_ROTINAAVALIAPRC     D  ")
  qd.Add("                         WHERE D.HANDLE = :HANDLE         ")
  qd.Add("                           AND C.AVALIAPRC = D.HANDLE )   ")
  qd.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qd.ExecSQL

  Set qd = Nothing

  If InTransaction Then
    Commit
  End If

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC")

End Sub


Public Sub BOTAOESTADO_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If


  Dim q2 As Object

  Set q2 = NewQuery

  q2.Add("SELECT HANDLE FROM SAM_ROTINAAVALIAPRC_EST WHERE AVALIAPRC = :HANDLEAVALIA")
  q2.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  q2.Active = True

  If Not q2.EOF Then
    bsShowMessage("Avaliação de Estado já Processada.", "I")
    Exit Sub
  End If

  Set q2 = Nothing

  Dim vmsgRetorno As String

  Set Obj = CreateBennerObject("BSINTERFACE0010.RotinaGerarAvaliacao")

  'Obj.AvaliaPrecoGerarEstados(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vmsgRetorno )
  Obj.AvaliaPrecoGerarEstado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "E", 0, 0, 0, 0, 0, 0, vmsgRetorno )' pedir confirmação

  If vmsgRetorno <> "" Then
    bsShowMessage(vmsgRetorno, "E")
  End If


  CurrentQuery.Active = False
  CurrentQuery.Active = True

  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_EST")

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

  If CheckParametros = False Then
    Exit Sub
  End If


  Dim vmsgRetorno As String


  Set Obj = CreateBennerObject("BSINTERFACE0010.RotinaGerarAvaliacao")

  'Obj.AvaliaPrecoGerarAvaliacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, vmsgRetorno )

  Obj.AvaliaPrecoGerarAvaliacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "G", 0, 0, 0, 0, 0, 0, vmsgRetorno ) ' pedir confirmação


  If vmsgRetorno <> "" Then
    bsShowMessage(vmsgRetorno, "E")
  End If


  CurrentQuery.Active = False
  CurrentQuery.Active = True

  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC")

End Sub



Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOFINAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
End Sub



Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTOINICIAL.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim q2 As Object

  Set q2 = NewQuery

  q2.Add("SELECT HANDLE FROM SAM_ROTINAAVALIAPRC_val WHERE AVALIAPRC = :HANDLEAVALIA")
  q2.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  q2.Active = True

  If Not q2.EOF Then
    bsShowMessage("Excluir primeiro as avaliações pelo botão.", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim q3 As Object

  Set q3 = NewQuery

  q3.Add("SELECT HANDLE FROM SAM_ROTINAAVALIAPRC_Est WHERE AVALIAPRC = :HANDLEAVALIA")
  q3.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  q3.Active = True

  If Not q3.EOF Then
    bsShowMessage("Excluir primeiro as avaliações do Estado pelo botão.", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("situacao").AsString = "P" Then
    bsShowMessage("Avaliação já Processada. Alteração não permitida.", "E")
    CanContinue = False
    Exit Sub
  End If

  Set qp = NewQuery

  qp.Add("SELECT HANDLE FROM SAM_ROTINAAVALIAPRC_EST WHERE AVALIAPRC=:HANDLE")
  qp.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qp.Active = True

  If Not qp.EOF Then
    bsShowMessage("Avaliação com Estados já Cadastrados. Alteração não permitida.", "E")
    CanContinue = False
    Exit Sub
  End If


  If CheckParametros = False Then
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOCANCELARESTADO"
			BOTAOCANCELARESTADO_OnClick
		Case "BOTAOESTADO"
			BOTAOESTADO_OnClick
		Case "BOTAOGERAL"
			BOTAOGERAL_OnClick
	End Select
End Sub
