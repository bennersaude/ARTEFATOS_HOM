'HASH: 6AE9EF70917F8FF1823EC7E2F126CB03

'#Uses "*ProcuraEvento"

Dim qParamAtend As Object


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text)
  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)

  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")
  Set qParamAtend = NewQuery

  qParamAtend.Add("SELECT FILTRARGRAUSVALIDOS FROM SAM_PARAMETROSATENDIMENTO")
 Set qParamAtend.Active = True

  If qParamAtend.FieldByName("FILTRARGRAUSVALIDOS").AsString = "S" And Not CurrentQuery.FieldByName("EVENTO").IsNull Then
   vCriterio = "HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
  Else
  	vCriterio = ""
  End If

  Set qParamAtend = Nothing

  vColunas = "GRAU|DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  vCampos = "Grau|Descrição|Graus Válidos"
  vHandle = interface.Exec(CurrentSystem, "SAM_GRAU", vColunas, 2, vCampos, vCriterio, "Selecionando Grau", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  Set interface = Nothing
End Sub

