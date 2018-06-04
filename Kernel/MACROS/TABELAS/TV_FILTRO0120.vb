'HASH: 39A8D9322E012D081E56F3B3461C805B
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If VisibleMode Then
		NOMEARQUIVO.Visible = False
		ROTAVISO.Visible = False
	ElseIf WebMode Then
		NOMEARQUIVO.Visible = True
		ROTAVISO.Visible = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQLTipoFat As Object
	Dim Interface As Object

	Set SQLTipoFat = NewQuery
	SQLTipoFat.Clear
	SQLTipoFat.Add("SELECT HANDLE, CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CStr(CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger))
	SQLTipoFat.Active = True

	If SQLTipoFat.FieldByName("CODIGO").AsInteger <> 660 Then
		bsShowMessage("O Tipo de Faturamento deve ser recolhimento de Contribuições Federais!", "E")
		CanContinue = False
		Exit Sub
	End If

	On Error GoTo erro

	If VisibleMode Then

		Set Interface =CreateBennerObject("BSINTERFACE0031.Rotinas")
    	Interface.ContrFedRecolhimento(CurrentSystem,CurrentQuery.FieldByName("ROTINAFIN").AsInteger)

    ElseIf WebMode Then
    	Dim vsAux As String

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		vsAux = Interface.ContrFedRecolhimento(CurrentSystem, CurrentQuery.FieldByName("ROTINAFIN").AsInteger, CurrentQuery.FieldByName("NOMEARQUIVO").AsString)

		If vsAux <> "" Then
			bsShowMessage(vsAux, "E")
			CanContinue = False
			Exit Sub
		End If

    End If

    Set Interface =Nothing

	Exit Sub

	erro:
		bsShowMessage(Error, "E")
		CanContinue = False
End Sub
