'HASH: 0E463A22F75FB30805F2E8D0ED782A0D
'Macro: SAM_PRECOFILIAL_MINMAX
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

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

'#Uses "*ProcuraGrau"
Public Sub GRAU_OnPopup(ShowPopup As Boolean)
	If ShowPopup = False Then
		Exit Sub
	End If

	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraGrau(GRAU.Text)

	If vHandle <> 0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("GRAU").Value = vHandle
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  	Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
    Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  	Dim Msg As String
    If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("VALORMAXIMO").AsFloat <CurrentQuery.FieldByName("VALORMINIMO").AsFloat Then
		bsShowMessage("O valor máximo deve ser maior que o mínimo", "E")
		CanContinue = False
		Exit Sub
	End If

	Dim Interface As Object
	Dim Linha As String
	Dim vCondicao As String

	vCondicao = "AND EVENTO = " + CurrentQuery.FieldByName("EVENTO").AsString
	vCondicao = vCondicao + " AND GRAU = " + CurrentQuery.FieldByName("GRAU").AsString

	If CurrentQuery.FieldByName("CONVENIO").IsNull Then
		vCondicao = vCondicao + " AND CONVENIO IS NULL"
	Else
		vCondicao = vCondicao + " AND CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString
	End If

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRECOFILIAL_MINMAX", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "FILIAL", vCondicao)

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
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

