'HASH: 18EFA18F552E1E7DD79B3CF4453BE422
'Macro: SAM_MODULO_EVENTO
'#Uses "*bsShowMessage"
Option Explicit

'#Uses "*ProcuraEvento"

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
	If WebMode Then
		If WebMenuCode = "T5674" Then
			EVENTO.ReadOnly = True
		End If
		If WebMenuCode = "T1611" Then
			MODULO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim qVerificaEvento As Object
  Set qVerificaEvento = NewQuery

  qVerificaEvento.Active = False
  qVerificaEvento.Clear
  qVerificaEvento.Add("SELECT Count(1) Encontrou ")
  qVerificaEvento.Add("  FROM SAM_CONTRATO_MODEVENTO ")
  qVerificaEvento.Add(" WHERE EVENTO = "+CurrentQuery.FieldByName("EVENTO").AsString )

  qVerificaEvento.Active = True

  If qVerificaEvento.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Não é possível a exclusão deste eventos pois existe relacionamento cadastrado no Evento do Módulo do Contrato!", "E")
    CanContinue = False
    Set qVerificaEvento = Nothing
    Exit Sub
  End If

  Set qVerificaEvento = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificaDuplicidade As Object
  Set qVerificaDuplicidade = NewQuery


  qVerificaDuplicidade.Active = False
  qVerificaDuplicidade.Clear
  qVerificaDuplicidade.Add("SELECT Count(1) Encontrou ")
  qVerificaDuplicidade.Add("  FROM SAM_MODULO_EVENTO ")
  qVerificaDuplicidade.Add(" WHERE HANDLE <> "+CurrentQuery.FieldByName("HANDLE").AsString  )
  qVerificaDuplicidade.Add("   AND EVENTO =  "+CurrentQuery.FieldByName("EVENTO").AsString )
  'qVerificaDuplicidade.Add("   AND MODULO = "+Str(RecordHandleOfTable("SAM_MODULO")) )

  qVerificaDuplicidade.Add("   AND MODULO = " + Str(CurrentQuery.FieldByName("MODULO").AsInteger))
  qVerificaDuplicidade.Active = True

  If qVerificaDuplicidade.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Evento já cadastrado!", "E")
    CanContinue = False
    Set qVerificaDuplicidade = Nothing
    Exit Sub
  End If

  Set qVerificaDuplicidade = Nothing

End Sub
