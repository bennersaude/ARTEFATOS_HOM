'HASH: 1E31FC7DC9EADFC72E3A7B58E1635E6D
'Macro: SAM_TABPF_EVENTO
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)

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
		ElseIf WebMenuCode = "T1244" Then
			TABELAPFEVENTO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim Consulta As Object
Set Consulta = NewQuery

Consulta.Clear
Consulta.Add("SELECT *                                ")
Consulta.Add("  FROM SAM_TABPF_EVENTO                 ")
Consulta.Add(" WHERE TABELAPFEVENTO = :TABPFEVENTO    ")
Consulta.Add("   AND EVENTO = :EVENTO                 ")
If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
  Consulta.Add(" AND HANDLE <> :HANDLE                ")
  Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
End If
Consulta.ParamByName("TABPFEVENTO").AsInteger = CurrentQuery.FieldByName("TABELAPFEVENTO").AsInteger
Consulta.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
Consulta.Active = True

If Not Consulta.FieldByName("HANDLE").IsNull Then
  bsShowMessage("Evento já cadastrado!", "E")
  CanContinue = False
  Exit Sub
End If
End Sub
