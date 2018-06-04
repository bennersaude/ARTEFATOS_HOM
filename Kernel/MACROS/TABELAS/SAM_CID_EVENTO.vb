'HASH: 5CE563B08E2B0801C37A5AB806B0665C

Public Sub BOTAOGERARCID_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_CID_EVENTO", "Duplicando Cids para o evento", "SAM_CID", "CID", "EVENTO", RecordHandleOfTable("SAM_TGE"), "P", "ESTRUTURA")
  Set Obj = Nothing
  CurrentQuery.Active = False
  CurrentQuery.Active = True
  CurrentQuery.First
  RefreshNodesWithTable("SAM_CID_EVENTO")


End Sub

Public Sub CID_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  If Not VisibleMode Then
    Exit Sub
  End If

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")
  vColunas = "ESTRUTURA|DESCRICAO"
  vCriterio = "ULTIMONIVEL = 'S' "
  vCampos = "Estrutura|Descricao"
  vHandle = interface.Exec(CurrentSystem, "SAM_CID", vColunas, 1, vCampos, vCriterio, "C.I.D.", False, "")
  If vHandle <> 0 Then
    CurrentQuery.FieldByName("CID").Value = vHandle
  End If
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterScroll()

  If (WebMode) Then
    CID.WebLocalWhere = "ULTIMONIVEL = 'S' "
  End If

End Sub
