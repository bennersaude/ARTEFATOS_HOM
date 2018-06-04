'HASH: B03795EB569903717FA55DC76E08CF22
'Macro: SAM_CONTRATO_CORRESPONS
'''

Public Sub CONTRATOCORRESPONS_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO|TABTIPOCONTRATO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND LOCALFATURAMENTO = 'F'"
  '  vCriterio =vCriterio + "AND HANDLE <> " +Str(CurrentQuery.FieldByName("CONTRATO").AsInteger)
  vCampos = "Contrato|Contratante|Data Adesão|Tipo de Contrato"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 2, vCampos, vCriterio, "Contratos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOCORRESPONS").Value = vHandle
  End If
  Set interface = Nothing
End Sub

