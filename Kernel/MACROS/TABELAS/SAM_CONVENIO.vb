'HASH: 854BA7B1A56DBA40E85C2D36CE8BD0D0
'MACRO: SAM_CONVENIO
'#Uses "*bsShowMessage"

Dim vCondicao As String
Dim Checagem As Long

Public Sub TABLE_AfterEdit()
	UpdateLastUpdate("SAM_CONVENIO")

	Dim sql As Object
	Set sql = NewQuery

	sql.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE OR HANDLE = :pHANDLE")

	sql.ParamByName("pHANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
	sql.Active = True

	vCondicao = ""
	If VisibleMode Then
		vCondicao = vCondicao + "SAM_CONVENIO.HANDLE "
	Else
		vCondicao = vCondicao + "A.HANDLE"
	End If
	vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = " + sql.FieldByName("HANDLE").AsString + ") "

	sql.Next

	While Not sql.EOF
		vCondicao = vCondicao + " OR "
		If VisibleMode Then
			vCondicao = vCondicao + "SAM_CONVENIO.HANDLE "
		Else
			vCondicao = vCondicao + "A.HANDLE"
		End If
		vCondicao = vCondicao + "IN (SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = " + sql.FieldByName("HANDLE").AsString + ") "

		sql.Next
	Wend

	If VisibleMode Then
		CONVENIOMESTRE.LocalWhere = vCondicao
	Else
		CONVENIOMESTRE.WebLocalWhere = vCondicao
	End If

	Set sql = Nothing
End Sub

Public Sub TABLE_AfterPost()
  If (Not WebMode) Then
	If Not CONVENIOMESTRE.Visible Then
		CurrentQuery.Edit
	End If
  Else
    If (CurrentQuery.FieldByName("CONVENIOMESTRE").IsNull) Then
      Dim qAux As Object
      Set qAux = NewQuery
      qAux.Clear
      qAux.Add("UPDATE SAM_CONVENIO ")
      qAux.Add("   SET CONVENIOMESTRE = :HANDLE ")
      qAux.Add(" WHERE HANDLE = :HANDLE ")
      qAux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qAux.ExecSQL
      Set qAux = Nothing
    End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If (Not WebMode) Then
	If CONVENIOMESTRE.Visible Then
		If CurrentQuery.FieldByName("CONVENIOMESTRE").IsNull Then
			bsShowMessage("Campo 'Utilizar preços de' é obrigatório !", "E")
			CanContinue = False
		End If
	Else
		bsShowMessage("Informar campo 'Utilizar preços de' ", "I")
	End If
  End If
End Sub

Public Sub TABLE_AfterScroll()
	CONVENIOMESTRE.Visible = True
End Sub

Public Sub TABLE_AfterInsert()
	CONVENIOMESTRE.Visible = False
End Sub

'Valeska sms 19128
Public Sub RELATORIOAVISOCANCELAMENTO_OnBtnClick()
	Dim OLEAutorizador As Object
	Dim handlexx As Long

	On Error GoTo cancel

	Set OLEAutorizador = CreateBennerObject("Procura.Procurar")

	handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório", "CODIGO = 'BEN003B'", "Procura por Relatórios", True, "")

	If handlexx <> 0 Then
		Dim SQL As Object
		Set SQL = NewQuery

		SQL.Add("SELECT CODIGO FROM R_RELATORIOS WHERE HANDLE = :HANDLE")

		SQL.ParamByName("HANDLE").Value = handlexx
		SQL.Active = True

		If CurrentQuery.State = 1 Then
			CurrentQuery.Edit
		End If

		CurrentQuery.FieldByName("RELATORIOAVISOCANCELAMENTO").Value = SQL.FieldByName("CODIGO").AsString
	End If

	Set OLEAutorizador = Nothing

	cancel :
End Sub
