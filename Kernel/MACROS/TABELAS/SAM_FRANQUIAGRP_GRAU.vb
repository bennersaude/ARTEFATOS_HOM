'HASH: 52589E4A69B10C69C22C80A149FDCD7E

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vColunas As String
  Dim vCriterios As String
  Dim vHandle As Long
  Dim vCampos As String
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO"

  vCriterios = ""
  vCampos = "Grau|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterios, "Tabela de Graus", False, "")
  CurrentQuery.FieldByName("GRAU").AsInteger = vHandle
  Set interface = Nothing
  ShowPopup = False


End Sub

