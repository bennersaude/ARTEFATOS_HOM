'HASH: F60546D10B472811B12AF7D1CCFC778D
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Confirme ou cancele processo, para ter sequência.", "I")
    Exit Sub
  End If

  Dim qp As Object

  Set qp = NewQuery

  qp.Add("SELECT HANDLE FROM SAM_ROTINAAVALIAPRC_PRE_VAL WHERE AVALIAPRC=:HANDLE")
  qp.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qp.Active = True


  If Not qp.EOF Then
    If bsShowMessage("Excluir Avaliação do Prestador?", "Q") = vbYes Then

      If Not InTransaction Then StartTransaction

      While Not qp.EOF

        Dim q1 As Object

        Set q1 = NewQuery

        q1.Add("DELETE FROM SAM_ROTINAAVALIAPRC_PRE_VAL WHERE AVALIAPRC = :HANDLEAVALIA")
        q1.ParamByName("HANDLEAVALIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        q1.ExecSQL

        Set q1 = Nothing

        qp.Next

      Wend

      Dim q2 As Object

      Set q2 = NewQuery

      q2.Add("UPDATE SAM_ROTINAAVALIAPRC_PRE SET SITUACAO = 'A' WHERE HANDLE = :HANDLEAVALIA")
      q2.ParamByName("HANDLEAVALIA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      q2.ExecSQL

      Commit

      RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_PRE")

      Set q2 = Nothing

    End If
  Else
    If Not InTransaction Then StartTransaction

    Dim dp As Object
    Set dp = NewQuery

    dp.Add("delete from SAM_ROTINAAVALIAPRC_PRE_VAL WHERE AVALIAPRC=:HANDLE")
    dp.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    dp.ExecSQL

    If InTransaction Then Commit

    '    MsgBox("Exclusão não efetuada. Avaliação Aberta.",vbOkOnly,"Avaliação de Preço")
    RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_PRE")

    Exit Sub
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

  'Set Obj = CreateBennerObject("SAMROTAVALIAPRC.geral")
  'Obj.executar(CurrentSystem, RecordHandleOfTable("SAM_ROTINAAVALIAPRC"), "P", 0, CurrentQuery.FieldByName("AVALIAPRC").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ESTADOS").AsInteger, CurrentQuery.FieldByName("MUNICIPIOS").AsInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)' pedir confirmação

  Dim vmsgRetorno As String
  Set Obj = CreateBennerObject("BSINTERFACE0010.RotinaGerarAvaliacao")
  Obj.AvaliaPrecoGerarAvaliacao(CurrentSystem, RecordHandleOfTable("SAM_ROTINAAVALIAPRC"), "P", 0, CurrentQuery.FieldByName("AVALIAPRC").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ESTADOS").AsInteger, CurrentQuery.FieldByName("MUNICIPIOS").AsInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)' pedir confirmação

  If vmsgRetorno <> "" Then
    bsShowMessage(vmsgRetorno, "E")
  End If


  CurrentQuery.Active = False
  CurrentQuery.Active = True

  Set Obj = Nothing

  RefreshNodesWithTable("SAM_ROTINAAVALIAPRC_PRE")

End Sub

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)

  Dim Mat As Object
  Dim vMatricula As Long
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False

  Set Mat = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_PRESTADOR.NOME|SAM_PRESTADOR.PRESTADOR"

  vCriterio = "MUNICIPIOPAGAMENTO = " + CurrentQuery.FieldByName("MUNICIPIOS").AsString

  vCampos = "Nome|CPF/CNPJ"

  vHandle = Mat.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Prestadores", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterScroll()
  If CurrentQuery.FieldByName("SITUACAO").AsString = "P" Then
    PRESTADOR.ReadOnly = True
  Else
    PRESTADOR.ReadOnly = False
  End If

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  bsShowMessage("Para excluir, deve-se usar o botão cancelar no processo acima.", "E")
  CanContinue = False
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOGERAL"
			BOTAOGERAL_OnClick
	End Select
End Sub
