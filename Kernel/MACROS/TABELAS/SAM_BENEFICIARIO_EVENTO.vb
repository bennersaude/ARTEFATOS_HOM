'HASH: DAB8FB1192D9DF1F8E6A665A838B52BA

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|DESCRICAO|ALTACOMPLEXIDADE"
  vCriterio = "ALTACOMPLEXIDADE = 'S' AND ULTIMONIVEL = 'S'"
  vCampos = "Estrutura|Descricao|Altacomplexidade"
  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterio, "Evento", False, EVENTO.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_BENEFICIARIO_EVENTO", "Duplicando Eventos para Beneficiário", "SAM_TGE", "EVENTO", "BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "SAM_BENEFICIARIO_EVENTO", "Excluindo Eventos para CID", "SAM_TGE", "EVENTO", "BENEFICIARIO", CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub

