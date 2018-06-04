'HASH: 5A1CCCC19D65B0E8A3078AD9B5DF288C
'Macro: SAM_MIGRACAO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub CONTRATODESTINO_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTexto As String
  Dim vTipopesquisa As Integer
  vTipopesquisa=2

  vTexto = CONTRATOORIGEM.LocateText

  If IsNumeric(vTexto) Then
    vTipopesquisa=1
  End If

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vTipopesquisa, vCampos, vCriterio, "Contratos", True, vTexto)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATODESTINO").Value = vHandle
  End If

  Set interface = Nothing

End Sub

Public Sub CONTRATOORIGEM_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTexto As String
  Dim vTipopesquisa As Integer
  vTipopesquisa=2

  vTexto = CONTRATOORIGEM.LocateText

  If IsNumeric(vTexto) Then
    vTipopesquisa=1
  End If


  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"
  vCriterio = "DATACANCELAMENTO IS NULL "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vTipopesquisa, vCampos, vCriterio, "Contratos", True, vTexto)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOORIGEM").Value = vHandle
  End If

  Set interface = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'CHECAR VIGENCIA
  Dim interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = "AND CONTRATODESTINO = " + CurrentQuery.FieldByName("CONTRATODESTINO").AsString

  Set interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = interface.Vigencia(CurrentSystem, "SAM_MIGRACAO", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "CONTRATOORIGEM", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If

End Sub

