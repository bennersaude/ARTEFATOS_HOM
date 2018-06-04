'HASH: 5C2078C3BDF38F1BE1462722563AA434
'#Uses "*bsShowMessage"

Option Explicit
Dim v_EventoInicial As Long
Dim v_EventoFinal As Long

Public Sub GRAUINICIAL_OnPopup(ShowPopup As Boolean)
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
  CurrentQuery.FieldByName("GRAUINICIAL").AsInteger = vHandle
  Set interface = Nothing
  ShowPopup = False

End Sub


Public Sub GRAUFINAL_OnPopup(ShowPopup As Boolean)
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
  CurrentQuery.FieldByName("GRAUFINAL").AsInteger = vHandle
  Set interface = Nothing
  ShowPopup = False

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q As Object
  Set Q = NewQuery

  Q.Active = False
  Q.Add("SELECT T1.ESTRUTURA GRAUINICIAL,")
  Q.Add("       T2.ESTRUTURA GRAUFINAL   ")
  Q.Add("  FROM SAM_TGE T1,                ")
  Q.Add("       SAM_TGE T2                 ")
  Q.Add(" WHERE T1.HANDLE = :GRAUINICIAL ")
  Q.Add("   AND T2.HANDLE = :GRAUFINAL   ")

  Q.ParamByName("GRAUINICIAL").Value = CurrentQuery.FieldByName("GRAUINICIAL").AsInteger
  Q.ParamByName("GRAUFINAL").Value = CurrentQuery.FieldByName("GRAUFINAL").AsInteger
  Q.Active = True

  If Q.FieldByName("GRAUINICIAL").AsString >Q.FieldByName("GRAUFINAL").AsString Then
    bsShowMessage("Grau Inicial não pode ser maior que Grau Final !", "E")
    CanContinue = False
  End If

  Set Q = Nothing
End Sub

