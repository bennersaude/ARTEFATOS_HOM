'HASH: 8EBCD0EECD0BF6A90B8A248E59CE9F89
'Macro: SAM_ALERTACONTRATO
'#Uses "*bsShowMessage"

Option Explicit
'#Uses "*ProcuraContrato"
Dim vgDataFinal As Date

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()
  If CurrentQuery.State = 3 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If
  Dim sql As Object
  Set sql = NewQuery

  If Not InTransaction Then StartTransaction

  sql.Add("UPDATE SAM_ALERTACONTRATO SET USUARIO=:USUARIO, DATA=:DATA WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
  sql.ParamByName("USUARIO").Value = CurrentUser
  sql.ParamByName("DATA").Value = ServerNow
  sql.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set sql = Nothing
End Sub

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraContrato(CONTRATO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATO").Value = vHandle
  End If


End Sub


Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub


Public Sub TABLE_AfterScroll()
  BOTAOGERAREVENTOS.Visible=False
  
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAFINAL.ReadOnly = False
    DESCRICAO.ReadOnly = False
    AUTORIZACAOACAO.ReadOnly = False
    AUTORIZACAOEXECUTOR.ReadOnly = False
    AUTORIZACAOSOLICITANTE.ReadOnly = False
    MOTIVONEGACAO.ReadOnly = False
    ACAOPAGAMENTO.ReadOnly = False
    PAGAMENTOEXECUTOR.ReadOnly = False
    PAGAMENTORECEBEDOR.ReadOnly = False
    MOTIVOGLOSA.ReadOnly = False
    ALERTATEXTO.ReadOnly = False
    GERAAUDITORIA.ReadOnly = False
    AUTORIZACAOLOCALEXEC.ReadOnly = False
    PAGAMENTOLOCALEXEC.ReadOnly = False
    AUTORIZACAORECEBEDOR.ReadOnly = False
    TITULAR.ReadOnly = False
  Else
    DATAFINAL.ReadOnly = True
    DESCRICAO.ReadOnly = True
    AUTORIZACAOACAO.ReadOnly = True
    AUTORIZACAOEXECUTOR.ReadOnly = True
    AUTORIZACAOSOLICITANTE.ReadOnly = True
    MOTIVONEGACAO.ReadOnly = True
    ACAOPAGAMENTO.ReadOnly = True
    PAGAMENTOEXECUTOR.ReadOnly = True
    PAGAMENTORECEBEDOR.ReadOnly = True
    MOTIVOGLOSA.ReadOnly = True
    ALERTATEXTO.ReadOnly = True
    GERAAUDITORIA.ReadOnly = True
    AUTORIZACAOLOCALEXEC.ReadOnly = True
    PAGAMENTOLOCALEXEC.ReadOnly = True
    AUTORIZACAORECEBEDOR.ReadOnly = True
    TITULAR.ReadOnly = True
  End If
  vgDataFinal = CurrentQuery.FieldByName("DATAFINAL").AsDateTime

  If WebMode Then
  	If WebMenuCode = "T2886" Then
  		CONTRATO.ReadOnly = True
  	End If
  End If
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
  Q.Add("DELETE FROM SAM_ALERTACONTRATO_EVENTO WHERE CONTRATOALERTA= :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '--------------------------------------------------------------------------------
  Q.Clear
  Q.Add("DELETE FROM SAM_ALERTACONTRATO_PRESTADOR WHERE CONTRATOALERTA = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '--------------------------------------------------------------------------------
  Q.Clear
  Q.Add("DELETE FROM SAM_ALERTACONTRATO_TPDEP WHERE ALERTACONTRATO = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '**************** Fim ***********************************************************

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  '  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
  '     CanContinue =False
  '     MsgBox("Registro finalizado não pode ser alterado!")
  '     Exit Sub
  '  End If

  If CurrentUser <>CurrentQuery.FieldByName("USUARIO").AsInteger Then
    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário diferente", "E")
    Exit Sub
  End If
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
