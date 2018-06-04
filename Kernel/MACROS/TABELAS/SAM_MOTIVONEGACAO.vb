'HASH: DFCF56F962A48CBF8F5C089834525482
'Macro: SAM_MOTIVONEGACAO

'#Uses "*bsShowMessage"


Public Sub MOTIVONEGACAO_OnChange()
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Clear
  q1.Add("SELECT DESCRICAO FROM SIS_MOTIVONEGACAO WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVONEGACAO").Value
  q1.Active = True
  CurrentQuery.Edit
  CurrentQuery.FieldByName("DESCRICAO").Value = q1.FieldByName("DESCRICAO").Value
  q1.Active = False
  Set q1 = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

'SMS 98903 - Ricardo Rocha - 03/07/2008
  If InStr(CurrentQuery.FieldByName("DESCRICAO").AsString, "<") > 0 Then
  	bsShowMessage("Caracter '<' não é válido na descrição. Favor utilizar outro caracter.", "E")
  	CanContinue = False
  	Exit Sub
  ElseIf InStr(CurrentQuery.FieldByName("DESCRICAO").AsString, ">") > 0 Then
	bsShowMessage("Caracter '>' não é válido na descrição. Favor utilizar outro caracter.", "E")
  	CanContinue = False
  	Exit Sub
  End If

End Sub
