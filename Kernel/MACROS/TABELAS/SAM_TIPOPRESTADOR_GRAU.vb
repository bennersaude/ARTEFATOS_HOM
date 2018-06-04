'HASH: C8180AAF3E73C69C743AFAA7F097C81A
'Macro: SAM_TIPOPRESTADOR_GRAU
'#Uses "*ProcuraGrau"
'#Uses "*bsShowMessage"


Public Sub GRAU_OnPopup(ShowPopup As Boolean)
  'If Len(GRAU.Text) = 0 Then
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraGrau(GRAU.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRAU").Value = vHandle
  End If
  ' End If
End Sub

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("GRAU").AsInteger = RecordHandleOfTable("SAM_GRAU")
End Sub


Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_SAM_TIPOPRESTADOR_GRAU_1432" Then
			TIPOPRESTADOR.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT HANDLE FROM SAM_TIPOPRESTADOR_GRAU WHERE TIPOPRESTADOR = :TIPOPRESTADOR AND GRAU = :GRAU AND HANDLE <> :HANDLE")
  Q.ParamByName("TIPOPRESTADOR").Value = CurrentQuery.FieldByName("TIPOPRESTADOR").Value
  Q.ParamByName("GRAU").Value = CurrentQuery.FieldByName("GRAU").Value
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  Q.Active = True
  If Not Q.EOF Then
    bsShowMessage("Grau já está cadastrado !", "E")
    CanContinue = False
    Set Q = Nothing
    Exit Sub
  End If
  Set Q = Nothing

  If CurrentQuery.FieldByName("GRAU").AsInteger = 0 Then
    bsShowMessage("É necessário informar o grau.", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

