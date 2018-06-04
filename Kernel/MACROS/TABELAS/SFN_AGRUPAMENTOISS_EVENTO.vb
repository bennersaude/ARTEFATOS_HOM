'HASH: 939ADEE3278A9147A8FA20291D6B4379
 

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "ESTRUTURA|ESTRUTURANUMERICA|DESCRICAO|ULTIMONIVEL"

  vCampos = "Estrutura|Estrutura Numerica|Descrição do Evento|Último Nível"

  vHandle = Interface.Exec(CurrentSystem, "SAM_TGE", vColunas, 2, vCampos, "", "Eventos", True, "")


  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

  Set Interface = Nothing


  
End Sub
