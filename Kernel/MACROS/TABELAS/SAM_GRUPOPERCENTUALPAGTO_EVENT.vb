'HASH: E809286D4866B72DE87EE02AFEF2A5E7
' sam_grupopercentualpagto_event

'Macro: sam_grupopercentualpagto_event
Option Explicit

'#Uses "*ProcuraEvento"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraEvento(True, EVENTO.Text)

  If vHandle<>0 Then
    Dim qBuscaEvento As Object
    Set qBuscaEvento = NewQuery

    qBuscaEvento.Clear
    qBuscaEvento.Add("SELECT CIRURGICO FROM SAM_TGE WHERE HANDLE = :HTGE")
    qBuscaEvento.ParamByName("HTGE").AsInteger = vHandle
    qBuscaEvento.Active = True

    If qBuscaEvento.FieldByName("CIRURGICO").AsString = "S" Then
      MsgBox("Não é possível cadastrar eventos cirúrgicos!")

      CurrentQuery.Edit
      CurrentQuery.FieldByName("EVENTO").Clear
    Else
      CurrentQuery.Edit
      CurrentQuery.FieldByName("EVENTO").Value = vHandle
    End If

    Set qBuscaEvento = Nothing
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qAux As Object
  Set qAux = NewQuery
  qAux.Clear
  qAux.Add("SELECT COUNT(1) QTD FROM SAM_GRUPOPERCENTUALPAGTO_EVENT WHERE EVENTO = :EVENTO AND GRUPOPERCENTUALPAGTO <> :GRUPOPERCENTUALPAGTO")
  qAux.ParamByName("GRUPOPERCENTUALPAGTO").Value = CurrentQuery.FieldByName("GRUPOPERCENTUALPAGTO").Value
  qAux.ParamByName("EVENTO").Value = CurrentQuery.FieldByName("EVENTO").Value
  qAux.Active = True
  If qAux.FieldByName("QTD").AsInteger > 0 Then
    MsgBox("Evento já cadastrado em outro grupo de percentual, favor verificar")
    CanContinue = False
    Set qAux = Nothing
    Exit Sub
  End If

  qAux.Clear
  qAux.Add("SELECT CIRURGICO FROM SAM_TGE WHERE HANDLE = :HTGE")
  qAux.ParamByName("HTGE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  qAux.Active = True
  If (qAux.FieldByName("CIRURGICO").AsString = "S") Then
    MsgBox("Não é possível cadastrar eventos cirúrgicos!")
    CanContinue = False
    Set qAux = Nothing
    Exit Sub
  End If

  Set qAux = Nothing
End Sub


