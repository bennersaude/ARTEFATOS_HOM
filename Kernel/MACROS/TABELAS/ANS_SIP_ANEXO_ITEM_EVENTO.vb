'HASH: 96578D0B4EC58DB1F4075FBD8BF0B880
'Macro: ANS_SIP_ANEXO_ITEM_EVENTO
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long

  ShowPopup = False

  vHandle = ProcuraEvento(True, EVENTO.Text)

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM ANS_SIP_ANEXO_ITEM_EVENTO WHERE SIPITEM = :SIPITEM AND EVENTO = :HEVENTO AND HANDLE <> :HANDLE")
  sql.ParamByName("SIPITEM").AsInteger = CurrentQuery.FieldByName("SIPITEM").AsInteger
  sql.ParamByName("HEVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True

  If Not sql.FieldByName("HANDLE").IsNull Then
  	bsShowMessage("Evento já cadastrado para esse ítem.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
