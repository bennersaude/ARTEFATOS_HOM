'HASH: BAA231F5C6F90A2F4D8BA365F499356D
'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Public Sub BOTAOGERAREVENTOS_OnClick()
	Dim Obj As Object
	Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")

	Obj.Gerar2(CurrentSystem, "SAM_PRESTADOR_AUTORIZEVENTO", "Eventos que necessitam de autorização", "SAM_TGE", _
			"EVENTO", "PRESTADOR", CurrentQuery.FieldByName("PRESTADOR").AsInteger, "S", "ESTRUTURA", 1, "REGIMEATENDIMENTO|", _
			CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsString)
	Set Obj = Nothing

	RefreshNodesWithTable("SAM_PRESTADOR_AUTORIZEVENTO")
End Sub

Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
	Dim vHandle As Long

	ShowPopup = False
	vHandle = ProcuraEvento(True, EVENTO.Text)

	If vHandle <>0 Then
		CurrentQuery.Edit
		CurrentQuery.FieldByName("EVENTO").Value = vHandle
	End If
End Sub

Public Sub TABLE_AfterScroll()

	If WebMode And _
	   CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger > 0 Then
       REGIMEATENDIMENTO.ReadOnly = True
    Else
       REGIMEATENDIMENTO.ReadOnly = False
	End If

End Sub


Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Clear
	SQL.Add("SELECT ISENTAAUTORIZ    ")
	SQL.Add("  FROM SAM_PRESTADOR    ")
	SQL.Add(" WHERE HANDLE = :HANDLE ")

	If (VisibleMode Or WebMode) Then
		SQL.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
  	Else
		SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
  	End If

	SQL.Active = True

	If (SQL.FieldByName("ISENTAAUTORIZ").AsString <> "S") Then

		If VisibleMode Then
			bsShowMessage("Prestador não está cadastrado como Isento de Autorização. Não é permitido incluir registros.", "E")
		Else
			CancelDescription = "Prestador não está cadastrado como Isento de Autorização. Não é permitido incluir registros."
		End If
		CanContinue = False
	End If

	Set SQL = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOGERAREVENTOS"
			BOTAOGERAREVENTOS_OnClick
	End Select
End Sub
