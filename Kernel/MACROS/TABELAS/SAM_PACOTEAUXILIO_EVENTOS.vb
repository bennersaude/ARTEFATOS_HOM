'HASH: 28A5E9827C9442E3B35FC260E0F66D25

'#Uses "*ProcuraEvento"
'#Uses "*ProcuraGrau"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(True, EVENTO.Text) ' só último nível
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
    CurrentQuery.FieldByName("GRAU").Clear 'Balani SMS 55210 23/12/2005
  End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  'Balani SMS 55210 23/12/2005
  Dim vHandle As Long
  ShowPopup = False
  'vHandle = ProcuraGrau

  Dim interface As Object
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String


  Set interface =CreateBennerObject("Procura.Procurar")

  vColunas ="SAM_GRAU.GRAU|SAM_GRAU.Z_DESCRICAO|SAM_GRAU.VERIFICAGRAUSVALIDOS"

  If CurrentQuery.FieldByName("EVENTO").IsNull Then
    vCriterio = ""
  Else
    vCriterio ="SAM_GRAU.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString + ")"
  End If

  vCampos ="Código do Grau|Descrição|Tipo do Grau|Graus Válidos"

  vHandle = interface.Exec(CurrentSystem,"SAM_GRAU",vColunas,2,vCampos,vCriterio,"Graus de Atuação",True,"")

  Set interface =Nothing
  If vHandle <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If

  'final SMS 55210
End Sub

