'HASH: 1F7446AC8EF1EFB337A0A117D9FDCA6A


Public Sub CID_OnEnter()
  CID.AnyLevel = True
End Sub

Public Sub CID_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|DESCRICAO"
  vCriterio = " ULTIMONIVEL = 'S'"
  vCampos = "Estrutura|Descricao"
  vHandle = interface.Exec(CurrentSystem, "SAM_CID", vColunas, 2, vCampos, vCriterio, "CID", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CID").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_BENEFICIARIO_CID", "Duplicando CID's para Beneficiário", "SAM_CID", "CID", "BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "SAM_BENEFICIARIO_CID", "Excluindo CID's do Beneficiário", "SAM_CID", "CID", "BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub

