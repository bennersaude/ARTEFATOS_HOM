'HASH: D38B00CB688433D6409259D88301E0A0
 
'PRE063

Option Explicit

Public Sub EVENTOFINAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO"

  vCriterio = "ULTIMONIVEL = 'S'  and INATIVO = 'N' and MASCARATGE = " & Str(CurrentQuery.FieldByName("MASCARATGE").AsInteger)
  vCampos = "Evento|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTOFINAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOFINAL").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub EVENTOINICIAL_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|SAM_TGE.DESCRICAO"

  vCriterio = "ULTIMONIVEL = 'S' and INATIVO = 'N' and MASCARATGE = " & Str(CurrentQuery.FieldByName("MASCARATGE").AsInteger)
  vCampos = "Evento|Descrição"

  vHandle = interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, vCriterio, "Tabela Geral de Eventos", True, EVENTOINICIAL.Text)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTOINICIAL").Value = vHandle
  End If
  Set interface = Nothing
End Sub

