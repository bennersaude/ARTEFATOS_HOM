'HASH: 897982C2D9357EE5FDF8A7CE43DDDDC1


Public Sub DEPENDENTEDESTINO_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim SelContrato As Object
  Dim vHandle As Long
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vTabela As String

  Set SelContrato = NewQuery

  SelContrato.Active = False
  SelContrato.Clear
  SelContrato.Add("SELECT CONTRATOMIGRACAO     ")
  SelContrato.Add("  FROM SAM_SITUACAORH         ")
  SelContrato.Add("  WHERE HANDLE = :HANDLE    ")
  SelContrato.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SITUACAORH").AsInteger
  SelContrato.Active = True


  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "SAM_TIPODEPENDENTE.DESCRICAO"
  vCriterio = "SAM_CONTRATO_TPDEP.CONTRATO = " + SelContrato.FieldByName("CONTRATOMIGRACAO").Value
  vCampos = "Dependente"

  vTabela = "SAM_CONTRATO_TPDEP|SAM_TIPODEPENDENTE[SAM_TIPODEPENDENTE.HANDLE = SAM_CONTRATO_TPDEP.TIPODEPENDENTE]"

  'Seleciona os modulos pertencentes ao contrato

  vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Tipo dependente", True, DEPENDENTEDESTINO.Text)

  If vHandle = 0 Then
    CurrentQuery.FieldByName("DEPENDENTEDESTINO").Value = Null
  Else
    CurrentQuery.FieldByName("DEPENDENTEDESTINO").Value = vHandle
  End If

  Set Interface = Nothing
  Set SelContrato = Nothing
End Sub

Public Sub DEPENDENTEORIGEM_OnPopup(ShowPopup As Boolean)
  ShowPopup = False

  Dim SelContrato As Object
  Dim vHandle As Long
  Dim Interface As Object
  Dim vColunas As String
  Dim vCriterio As String
  Dim vCampos As String
  Dim vTabela As String

  Set SelContrato = NewQuery

  SelContrato.Active = False
  SelContrato.Clear
  SelContrato.Add("SELECT CONTRATO ")
  SelContrato.Add("  FROM SAM_SITUACAORH   ")
  SelContrato.Add("  WHERE HANDLE = :HANDLE                      ")
  SelContrato.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("SITUACAORH").AsInteger
  SelContrato.Active = True


  Set Interface = CreateBennerObject("Procura.Procurar")
  vColunas = "SAM_TIPODEPENDENTE.DESCRICAO"
  vCriterio = "SAM_CONTRATO_TPDEP.CONTRATO = " + SelContrato.FieldByName("CONTRATO").Value + " AND SAM_CONTRATO_TPDEP.HANDLE NOT IN (SELECT DEPENDENTEORIGEM FROM SAM_SITUACAORH_TPDEP WHERE SITUACAORH = " + Str(CurrentQuery.FieldByName("SITUACAORH").AsInteger) + " ) "
  vCampos = "Dependente"

  vTabela = "SAM_CONTRATO_TPDEP|SAM_TIPODEPENDENTE[SAM_TIPODEPENDENTE.HANDLE = SAM_CONTRATO_TPDEP.TIPODEPENDENTE]"

  'Seleciona os modulos pertencentes ao contrato

  vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 1, vCampos, vCriterio, "Módulo", True, DEPENDENTEORIGEM.Text)

  If vHandle = 0 Then
    CurrentQuery.FieldByName("DEPENDENTEORIGEM").Value = Null
  Else
    CurrentQuery.FieldByName("DEPENDENTEORIGEM").Value = vHandle
  End If

  Set Interface = Nothing
  Set SelContrato = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CurrentQuery.FieldByName("SITUACAOORIGEM").AsInteger = CurrentQuery.FieldByName("SITUACAORH").AsInteger
End Sub

