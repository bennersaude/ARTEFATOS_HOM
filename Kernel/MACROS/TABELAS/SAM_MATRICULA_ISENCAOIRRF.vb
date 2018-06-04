'HASH: 8A07376DCC247517108A2634589C8175
'MACRO SAM_MATRICULA_ISENCAOIRRF

Public Sub MOTIVO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTabela  As String
  Dim vTitulo As String

  If MOTIVO.PopupCase <> 0 Then
    ShowPopup = False
    Set Interface = CreateBennerObject("Procura.Procurar")

    vCampos = "Código|Descrição"
    vColunas = "CODIGO|DESCRICAO"
    vTabela = "SAM_MOTIVOISENCAOIRRF"
    vTitulo = "Motivos de isenção de IRRF"
    vCriterio = ""
    vHandle = Interface.Exec(CurrentSystem, vTabela, vColunas, 2, vCampos, vCriterio, vTitulo, True, MOTIVO.LocateText)

    Set Interface = Nothing
  Else
    ShowPopup = True
  End If

  If vHandle > 0 Then
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("MOTIVO").AsInteger = vHandle
  End If
End Sub
