'HASH: 43F5DD209E2D76D8CF970F54CEDFE81A
'Macro: SAM_TGE_DEPENDENTE
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
  Dim interface As Object
  Set interface = CreateBennerObject("TGE.Rotinas")
  If interface.checkEventoDep(CurrentSystem, CurrentQuery.FieldByName("EVENTODEPENDENTE").Value) = True Then
    bsShowMessage("Evento não pode ser cadastrado, para evitar a recursividade (loop)", "E")
    CurrentQuery.Delete
    CanContinue = False
  End If

End Sub

Public Sub EVENTODEPENDENTE_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO|NIVELAUTORIZACAO"

  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Evento|Descrição|Nível"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTODEPENDENTE").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    EVENTODEPENDENTE.WebLocalWhere = " A.ULTIMONIVEL = 'S' "
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTODEPENDENTE").AsInteger Then
    bsShowMessage("Evento não pode estar contido nele mesmo, para evitar a recursividade (loop)", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

