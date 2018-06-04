'HASH: 56D43809B628D812C668C8BF3B971824

'SAM_PATOLOGIA_CID
'Criado por: Milton
'SMS 11535

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
  Obj.Gerar(CurrentSystem, "SAM_PATOLOGIA_CID", "Duplicando CID's para Patologias", "SAM_CID", "CID", "PATOLOGIA", CurrentQuery.FieldByName("PATOLOGIA").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "SAM_PATOLOGIA_CID", "Excluindo CID's da Patologias", "SAM_CID", "CID", "PATOLOGIA", CurrentQuery.FieldByName("PATOLOGIA").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub

