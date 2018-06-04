'HASH: 78C472A84FAA965F985FBBBA3148A4E9

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO"

  vCriterio = "ULTIMONIVEL = 'S' AND INATIVO = 'N'"
  vCampos = "Estrutura|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 1, vCampos, vCriterio, "Tabela Geral de Eventos", True, "")

  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").AsInteger = vHandle
  End If
  Set interface = Nothing
End Sub
