'HASH: 0DDCF2274C81F4EF83A64413D1B170C3
'Macro: SAM_PRECOGERAL_MINMAX
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"
'#Uses "*ProcuraGrau"

Option Explicit

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	If ShowPopup = False Then
		Exit Sub
	End If

	If Len(GRAU.Text) = 0 Then
		Dim vHandle As Long

		ShowPopup = False
		vHandle = ProcuraGrau(GRAU.Text)

		If vHandle <> 0 Then
			CurrentQuery.Edit
			CurrentQuery.FieldByName("GRAU").Value = vHandle
		End If
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T4377" Then
			CONVENIO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim vCondicao As String

	If CurrentQuery.FieldByName("VALORMAXIMO").AsFloat <CurrentQuery.FieldByName("VALORMINIMO").AsFloat Then
		bsShowMessage("O valor máximo deve ser maior que o mínimo", "E")
		CanContinue = False
		Exit Sub
	End If

	Dim INTERFACE As Object
	Dim Linha As String
	Set INTERFACE = CreateBennerObject("SAMGERAL.Vigencia")

	vCondicao = "AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	vCondicao = vCondicao + " AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		vCondicao = vCondicao + " AND CONVENIO IS NULL"
	Else
		vCondicao = vCondicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Linha = INTERFACE.Vigencia(CurrentSystem, "SAM_PRECOGERAL_MINMAX", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "EVENTO", vCondicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set INTERFACE = Nothing
End Sub

Public Sub TABLE_AfterInsert()
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT COUNT(*) TOTAL FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

	SQL.Active = True

	If SQL.FieldByName("TOTAL").AsInteger = 1 Then
		SQL.Active = False

		SQL.Clear

		SQL.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE CONVENIOMESTRE = HANDLE")

		SQL.Active = True

		CurrentQuery.FieldByName("CONVENIO").Value = SQL.FieldByName("HANDLE").Value
	End If

	Set SQL = Nothing
End Sub
