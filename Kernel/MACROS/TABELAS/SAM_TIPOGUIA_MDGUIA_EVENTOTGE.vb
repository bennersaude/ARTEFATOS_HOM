'HASH: 52C0FEE05D5C43919B7332CAEB138635
'SAM_TIPOGUIA_MDGUIA_EVENTOTGE

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
  Dim qEvento As BPesquisa
  Set qEvento = NewQuery

  qEvento.Clear
  qEvento.Add("  SELECT STME.HANDLE                          ")
  qEvento.Add("    FROM SAM_TIPOGUIA_MDGUIA_EVENTOTGE  STME  ")
  qEvento.Add("   WHERE STME.MODELOGUIA = :MODELOGUIA        ")
  qEvento.Add("     AND STME.EVENTO = :EVENTO                ")
  qEvento.Add("     AND STME.HANDLE <> :HANDLE               ")
  qEvento.ParamByName("MODELOGUIA").AsInteger = CurrentQuery.FieldByName("MODELOGUIA").AsInteger
  qEvento.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  qEvento.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qEvento.Active = True

  If Not qEvento.EOF Then
    bsshowmessage("O evento já existe neste modelo de Guia", "E")
    CanContinue = False
  End If

End Sub
