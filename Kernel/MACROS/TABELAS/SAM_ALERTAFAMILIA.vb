'HASH: 4DA539946D06E5EF5B3C47CE29CDC5C7
'Macro: SAM_ALERTAFAMILIA
Option Explicit

Dim vgDataFinal As Date

'#Uses "*ProcuraFamilia"
'#Uses "*bsShowMessage"

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()
  If CurrentQuery.State = 3 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If
  Dim sql As Object
  Set sql = NewQuery

  If Not InTransaction Then StartTransaction

  sql.Add("UPDATE SAM_ALERTAFAMILIA")
  sql.Add("   SET USUARIO = :USUARIO,")
  sql.Add("       DATA     = :DATA")
  sql.Add(" WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ParamByName("USUARIO").Value = CurrentUser
  sql.ParamByName("DATA").Value = ServerNow
  sql.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set sql = Nothing
End Sub

Public Sub FAMILIA_OnChange()
  MostraResponsavel
End Sub

Public Sub FAMILIA_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraFamilia(FAMILIA.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("FAMILIA").Value = vHandle
  End If
  MostraResponsavel
End Sub


Public Sub FAMILIA_OnSearch(InternalCode As Long)
  MostraResponsavel
End Sub

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub


Public Sub TABLE_AfterScroll()
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

  MostraResponsavel
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
  Q.Add("DELETE FROM SAM_ALERTAFAMILIA_EVENTO WHERE FAMILIAALERTA= :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '--------------------------------------------------------------------------------
  Q.Clear
  Q.Add("DELETE FROM SAM_ALERTAFAMILIA_PRESTADOR WHERE FAMILIAALERTA = :HANDLE")
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

Public Sub MostraResponsavel
  Dim sql As Object
  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT B.NOME RESP,'B' T")
  sql.Add("  FROM SAM_FAMILIA F,")
  sql.Add("       SAM_BENEFICIARIO B")
  sql.Add(" WHERE F.HANDLE = :F")
  sql.Add("   AND TABRESPONSAVEL = 1")
  sql.Add("   AND B.HANDLE = F.TITULARRESPONSAVEL")
  sql.Add("UNION")
  sql.Add("SELECT P.NOME RESP, 'P' T")
  sql.Add("  FROM SAM_FAMILIA F,")
  sql.Add("       SFN_PESSOA P")
  sql.Add(" WHERE F.HANDLE = :F")
  sql.Add("   AND TABRESPONSAVEL = 2")
  sql.Add("   AND P.HANDLE = F.PESSOARESPONSAVEL")
  sql.ParamByName("F").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
  sql.Active = True
  RESPONSAVEL.Text = "Resp: " + sql.FieldByName("RESP").AsString + " [" + _
                     IIf(sql.FieldByName("T").AsString = "B", "Beneficiário", "Pessoa") + "]"
  sql.Active = False
  Set sql = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
      Case "BOTAOALTERARRESPONSAVEL"
        BOTAOALTERARRESPONSAVEL_OnClick
  End Select
End Sub
