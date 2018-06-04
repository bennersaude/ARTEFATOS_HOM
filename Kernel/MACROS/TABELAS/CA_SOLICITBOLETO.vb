'HASH: 27C9079B73C1079B3372BA4BEAA90CA7

'############## CENTRAL DE ATENDIMENTO #################

Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim vCriterios As String
  Dim vCampos As String
  Dim Interface As Variant
  Set interface = CreateBennerObject("Procura.Procurar")
  vColunas = "BENEFICIARIO|NOME"
  vCriterios = ""
  vCampos = "Beneficiário|Nome"
  vHandle = interface.exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 2, vCampos, vCriterios, "Tabela de beneficiários", False, "")
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
  End If
  ShowPopup = False
  Set Interface = Nothing
End Sub


Public Sub BOTAOCANCELAR_OnClick()
  ' +++++++++dentro da dll da a mensagem de confirmacao
  If MsgBox("Confirma o cancelamento da solicitação?", vbYesNo) = vbNo Then
    Exit Sub
  End If

  Dim vDll As Object
  Dim vRetorno As Boolean

  Set vDll = CreateBennerObject("CA023.BOLETO")

  ' "CA_SOLICITBOLETO",

  Retorno = vDll.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  If Retorno = False Then
    Exit Sub
  Else
    WriteAudit("C", HandleOfTable("CA_SOLICITBOLETO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Solicitação seg. via de boleto - Cancelamento")
    RefreshNodesWithTable("CA_SOLICITBOLETO")
  End If

End Sub

Public Sub TABLE_AfterScroll()
  Select Case CurrentQuery.FieldByName("SITUACAO").AsString
    Case "C"
      BOTAOCANCELAR.Visible = False
    Case "P"
      BOTAOCANCELAR.Visible = False
    Case Else
      BOTAOCANCELAR.Visible = True
  End Select
End Sub

Public Sub TABLE_NewRecord()
  Dim vANO As String
  Dim SEQUENCIA As Long
  vANO = Format("yyyy", ServerDate)
  NewCounter("CA_ATEND", CDate(vANO), 1, SEQUENCIA)
  CurrentQuery.FieldByName("ANO").Value = ("01/01/" + vANO)
  CurrentQuery.FieldByName("NUMERO").Value = SEQUENCIA
End Sub


'###############################################################
