'HASH: A85B46114576F5A9B51D8F219F699F54
'Macro: SAM_ALERTAMATRICULA
'#Uses "*bsShowMessage"

Dim vgDataFinal As Date

Public Sub TABLE_AfterPost()
  TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
  MATRICULA.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull 'False
  DATAINICIAL.ReadOnly = Not CurrentQuery.FieldByName("DATAFINAL").IsNull
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

Public Sub BOTAOALTERARRESPONSAVEL_OnClick()

  If CurrentQuery.State = 3 Then
    bsShowMessage("O registro não pode estar em edição", "E")
    Exit Sub
  End If
  Dim SQL As Object
  Set SQL = NewQuery

  If Not InTransaction Then StartTransaction

  SQL.Add("UPDATE SAM_ALERTAMATRICULA SET USUARIO=:USUARIO, DATA=:DATA WHERE HANDLE=" + CurrentQuery.FieldByName("HANDLE").AsString)
  SQL.ParamByName("USUARIO").Value = CurrentUser
  SQL.ParamByName("DATA").Value = ServerNow
  SQL.ExecSQL

  If InTransaction Then Commit

  CurrentQuery.Active = False
  CurrentQuery.Active = True
  Set SQL = Nothing
End Sub

Public Sub MATRICULA_OnPopup(ShowPopup As Boolean)
  Dim Mat As Object
  Dim vMatricula As Long
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim SQL As Object

  ShowPopup = False

  Set Mat = CreateBennerObject("Procura.Procurar")
  vColunas = "SAM_MATRICULA.MATRICULA|SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_MATRICULA.Z_NOME|SAM_MATRICULA.INICIAIS|SAM_MATRICULA.NOMEMAE|SAM_MATRICULA.DATANASCIMENTO"
  vCriterio = "SAM_MATRICULA.DATAFALECIMENTO IS NULL "
  vCampos = "Matrícula|Matrícula Funcional|Nome|Iniciais|Nome da Mãe|Dt.Nascimento"
  vHandle = Mat.Exec(CurrentSystem, "SAM_MATRICULA|SAM_BENEFICIARIO[SAM_BENEFICIARIO.MATRICULA=SAM_MATRICULA.HANDLE]", vColunas, 3, vCampos, vCriterio, "Matrícula Única", False, "")
  If vHandle <>0 Then
    CurrentQuery.FieldByName("MATRICULA").Value = vHandle
  End If
  Set Mat = Nothing

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
  Q.Add("DELETE FROM SAM_ALERTAMATRICULA_EVENTO WHERE ALERTAMATRICULA= :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Q.ExecSQL
  '--------------------------------------------------------------------------------
  Q.Clear
  Q.Add("DELETE FROM SAM_ALERTAMATRICULA_PRESTADOR WHERE ALERTAMATRICULA = :HANDLE")
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
