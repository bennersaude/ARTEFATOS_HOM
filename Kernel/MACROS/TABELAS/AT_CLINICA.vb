'HASH: 7F880AB9A45671210DBDFA2D6871377E

'MACRO: AT_CLINICA

Public Sub PRESTADOR_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "NOME|PRESTADOR"

  vCriterio = " RECEBEDOR='S' and NAOFATURARGUIAS='S' "
  vCampos = "Nome|CNPJ/CPF"

  vHandle = interface.Exec(CurrentSystem, "SAM_PRESTADOR", vColunas, 1, vCampos, vCriterio, "Clínicas Próprias", False, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PRESTADOR").Value = vHandle
  End If
  Set interface = Nothing



End Sub

