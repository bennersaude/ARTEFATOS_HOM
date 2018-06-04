'HASH: B45F6CB3B014ABC4FC3F73B00A40B97E
'Macro: SAM_TGE_MUNICIPIO
'Última alteração: 07/05/2002
'Por: Milton - SMS 9017

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


Public Sub TABLE_AfterScroll()
  If (WebMode) Then
    EVENTO.WebLocalWhere = " A.ULTIMONIVEL = 'S' "
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim q1 As Object
  Set q1 = NewQuery

  q1.Clear
  q1.Active = False
  q1.Add("SELECT EVENTO, EVENTOMUNICIPIO FROM SAM_TGE_MUNICIPIO")
  'q1.Add("WHERE MUNICIPIO=:pMUNICIPIO AND (EVENTO=:pEVENTO OR EVENTOMUNICIPIO=:pEVENTOMUNICIPIO) AND HANDLE<>:pHANDLE")
  q1.Add("WHERE MUNICIPIO=:pMUNICIPIO AND EVENTOMUNICIPIO=:pEVENTOMUNICIPIO AND HANDLE<>:pHANDLE")
  q1.ParamByName("pHANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.ParamByName("pMUNICIPIO").Value = CurrentQuery.FieldByName("MUNICIPIO").Value
  'q1.ParamByName("pEVENTO").Value= CurrentQuery.FieldByName("EVENTO").Value
  q1.ParamByName("pEVENTOMUNICIPIO").Value = CurrentQuery.FieldByName("EVENTOMUNICIPIO").Value
  q1.Active = True

  If Not q1.EOF Then
    CanContinue = False
    bsShowMessage("Registro já cadastrado no Município", "E")
  End If
  Set q1 = Nothing
End Sub


