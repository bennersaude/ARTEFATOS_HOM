'HASH: 33670609F791EF8D09EDC705807E309C
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

 	Dim Interface As Object

	Dim SQLTipoFat As Object

	On Error GoTo erro

	Set SQLTipoFat = NewQuery

    SQLTipoFat.Clear

    SQLTipoFat.Add("SELECT HANDLE, CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE=" +CStr(CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger))

    SQLTipoFat.Active =True

    If (SQLTipoFat.FieldByName("CODIGO").AsInteger <> 640) Then
	  bSShowMessage("Tipo de Faturamento deve ser '640 - Recolhimento de ISS' !", "E")
	  CanContinue = False
      Exit Sub
	End If

	Set SQLTipoFat = Nothing


 	If VisibleMode Then

		Set Interface =CreateBennerObject("BSINTERFACE0031.Rotinas")
    	Interface.ISSRecolhimento(CurrentSystem,CurrentQuery.FieldByName("ROTINAFIN").AsInteger)

    ElseIf WebMode Then

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		Interface.ISSRecolhimento(CurrentSystem,CurrentQuery.FieldByName("ROTINAFIN").AsInteger, CurrentQuery.FieldByName("NOMEARQUIVO").AsString)

    End If

    Set Interface =Nothing

	Exit Sub

	erro:
		bsShowMessage(Error, "E")
		CanContinue = False
End Sub
