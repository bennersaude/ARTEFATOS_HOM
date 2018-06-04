'HASH: B592BF4DA0D0C2A250E86FBBDBFE3158
'Macro: SAM_TGE_ESTADO
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
  q1.Add("SELECT EVENTO, EVENTOESTADO FROM SAM_TGE_ESTADO")
  'q1.Add("WHERE ESTADO=:pESTADO AND (EVENTO=:pEVENTO OR EVENTOESTADO=:pEVENTOESTADO) AND HANDLE<>:pHANDLE")
  q1.Add("WHERE ESTADO=:pESTADO AND EVENTOESTADO=:pEVENTOESTADO AND HANDLE<>:pHANDLE")
  q1.ParamByName("pHANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  q1.ParamByName("pESTADO").Value = CurrentQuery.FieldByName("ESTADO").Value
  'q1.ParamByName("pEVENTO").Value= CurrentQuery.FieldByName("EVENTO").Value
  q1.ParamByName("pEVENTOESTADO").Value = CurrentQuery.FieldByName("EVENTOESTADO").Value
  q1.Active = True

  If Not q1.EOF Then
    CanContinue = False
    bsShowMessage("Registro já cadastrado no Estado", "E")
  End If
  Set q1 = Nothing

End Sub




