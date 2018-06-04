﻿'HASH: 5F1B8491F01AD68D09E978291118DC67
'#Uses "*bsShowMessage"

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vColunas As String
  Dim vCriterios As String
  Dim vHandle As Long
  Dim vCampos As String
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_GRAU.GRAU|SAM_GRAU.DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  vCriterios = ""
  vCampos = "Grau|Descrição|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterios, "Tabela de Graus", False, "")
  CurrentQuery.FieldByName("GRAU").AsInteger = vHandle
  Set interface = Nothing
  ShowPopup = False


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_TABPF_GRAU                   ")
Consulta.Add(" WHERE TABPFEVENTO = :TABPFEVENTO       ")
Consulta.Add("   AND GRAU = :GRAU                     ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("TABPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABPFEVENTO").AsInteger
Consulta.ParamByName("GRAU").AsInteger = CurrentQuery.FieldByName("GRAU").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Grau já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
