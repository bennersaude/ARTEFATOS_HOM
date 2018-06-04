'HASH: 0E5DA092132DD4C238FB8E092A58C33C
'Macro: SAM_MOTIVOGLOSA

'#Uses "*bsShowMessage"

Option Explicit

Public Sub MOTIVOGLOSA_OnChange()
  Dim q1 As Object
  Set q1 = NewQuery
  q1.Clear
  q1.Add("SELECT DESCRICAO FROM SIS_MOTIVOGLOSA WHERE HANDLE=:HANDLE")
  q1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("MOTIVOGLOSA").Value
  q1.Active = True
  CurrentQuery.Edit
  CurrentQuery.FieldByName("DESCRICAO").Value = q1.FieldByName("DESCRICAO").Value
  q1.Active = False
  Set q1 = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  Dim qParametrosProcContas As BPesquisa
  Set qParametrosProcContas = NewQuery

  qParametrosProcContas.Add("SELECT TABEMAILGLOSA FROM SAM_PARAMETROSPROCCONTAS")
  qParametrosProcContas.Active = True

  If (qParametrosProcContas.FieldByName("TABEMAILGLOSA").AsInteger = 1) Then
    COMUNICAPRESTFECHAMENTOAGRUP.Visible = False
  Else
    COMUNICAPRESTFECHAMENTOAGRUP.Visible = True
  End If

  qParametrosProcContas.Active = False
  Set qParametrosProcContas = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object
  Set sql = NewQuery
  sql.Add("SELECT HANDLE FROM SAM_MOTIVOGLOSA WHERE CODIGOGLOSA=:CODIGO AND HANDLE<>:HANDLE")
  sql.ParamByName("CODIGO").Value = CurrentQuery.FieldByName("CODIGOGLOSA").AsInteger
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active = True
  If Not sql.EOF Then
    bsShowMessage("Este código de glosa já foi cadastrado!", "E")
    CanContinue = False
    Exit Sub
  End If
  sql.Active = False
  Set sql = Nothing

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

