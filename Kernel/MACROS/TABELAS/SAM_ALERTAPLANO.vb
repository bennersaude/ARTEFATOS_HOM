'HASH: BBD3ECEBD20ACD19894FAD7767081897
'Macro: SAM_ALERTAPLANO
Option Explicit

Dim vgDataFinal As Date

'#Uses "*ProcuraPlano"
'#Uses "*bsShowMessage"

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()
  If CurrentQuery.State = 3 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If
  Dim sql As Object
  Set sql = NewQuery

  If Not InTransaction Then StartTransaction

  sql.Add("UPDATE SAM_ALERTAPLANO SET USUARIO=:USUARIO, DATA=:DATA WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
  sql.ParamByName("USUARIO").Value = CurrentUser
  sql.ParamByName("DATA").Value = ServerNow
  sql.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set sql = Nothing
End Sub

Public Sub PLANO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraPlano(PLANO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANO").Value = vHandle
  End If


End Sub


Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  DATAFINAL.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  DESCRICAO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  AUTORIZACAOACAO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  AUTORIZACAOEXECUTOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  AUTORIZACAOSOLICITANTE.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  MOTIVONEGACAO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  ACAOPAGAMENTO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  PAGAMENTOEXECUTOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  PAGAMENTORECEBEDOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  MOTIVOGLOSA.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  ALERTATEXTO.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  GERAAUDITORIA.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  AUTORIZACAOLOCALEXEC.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  PAGAMENTOLOCALEXEC.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
  AUTORIZACAORECEBEDOR.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário diferente", "E")
    Exit Sub
  End If

  '***************** SMS **********************************************************
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("DELETE FROM SAM_ALERTAPLANO_EVENTO WHERE PLANOALERTA = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '--------------------------------------------------------------------------------
  Q.Clear
  Q.Add("DELETE FROM SAM_ALERTAPLANO_PRESTADOR WHERE PLANOALERTA = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '**************** Fim ***********************************************************

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário diferente", "E")
    Exit Sub
  End If

  vgDataFinal = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vDllRTF2TXT As Object

  If(Not CurrentQuery.FieldByName("DATAFINAL").IsNull)And _
     (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)Then
  CanContinue = False
  bsShowMessage("A Data final, se informada, deve ser maior ou igual a inicial", "E")
  Exit Sub
End If

If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "N" And _
                            CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "N" Then
  CanContinue = False
  bsShowMessage("Pelo menos uma ação deve diferente de Nada", "E")
  Exit Sub
End If

If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString <>"N" And _
                             CurrentQuery.FieldByName("AUTORIZACAOEXECUTOR").AsString = "N" And _
                             CurrentQuery.FieldByName("AUTORIZACAOSOLICITANTE").AsString = "N" Then
  CanContinue = False
  bsShowMessage("Executor e/ou Solicitante para autorização deve ser selecionado", "E")
  Exit Sub
End If

If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString <>"N" And _
                             CurrentQuery.FieldByName("PAGAMENTOEXECUTOR").AsString = "N" And _
                             CurrentQuery.FieldByName("PAGAMENTORECEBEDOR").AsString = "N" Then
  CanContinue = False
  bsShowMessage("Executor e/ou Recebedor para pagamento deve ser selecionado", "E")
  Exit Sub
End If

If CurrentQuery.FieldByName("AUTORIZACAOACAO").AsString = "R" And _
                            CurrentQuery.FieldByName("MOTIVONEGACAO").IsNull Then
  CanContinue = False
  bsShowMessage("Para alerta de restrição na autorização deve ser informado o motivo de negação", "E")
  Exit Sub
End If

If CurrentQuery.FieldByName("ACAOPAGAMENTO").AsString = "R" And _
                            CurrentQuery.FieldByName("MOTIVOGLOSA").IsNull Then
  CanContinue = False
  bsShowMessage("Para alerta de restrição no pagamento deve ser informado o motivo de glosa", "E")
  Exit Sub
End If

Set vDllRTF2TXT = CreateBennerObject("RTF2TXT.Rotinas")
CurrentQuery.FieldByName("ALERTATEXTOTXT").AsString = vDllRTF2TXT.Rtf2Txt(CurrentSystem, CurrentQuery.FieldByName("ALERTATEXTO").AsString)
'SMS 59169 - Marcelo Barbosa - 15/03/2006
If InStr(CurrentQuery.FieldByName("ALERTATEXTOTXT").AsString,"{") > 0 Or _
   InStr(CurrentQuery.FieldByName("ALERTATEXTOTXT").AsString,"}") > 0 Then
     bsShowMessage("Não é permitido inserir os caracteres { (abre chave) e/ou } (fecha chave) no texto do Alerta!", "E")
     CanContinue = False
     Exit Sub
End If
'Fim - SMS 59169
Set vDllRTF2TXT = Nothing

If vgDataFinal <>CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
  If VisibleMode Then
    If bsShowMessage("Fechando a vigência não será permitido alteração no alerta , nem reabriar a vigência." + (Chr(13)) + _
                     "Deseja continuar?", "Q") = vbNo Then
      CanContinue = False
      Exit Sub
    End If
  Else
    bsShowMessage("A vigência foi fechada. Não será permitida a alteração do alerta!", "I")
  End If
End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
      Case "BOTAOALTERARRESPONSAVEL"
        BOTAOALTERARRESPONSAVEL_OnClick
  End Select
End Sub
