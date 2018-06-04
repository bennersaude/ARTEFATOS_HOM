'HASH: 16F996FABA900C7EBBD87EDE3079D75E
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

    If((SQLTipoFat.FieldByName("CODIGO").AsInteger)<>630)Then
		bsShowMessage("O TIPO DE FATURAMENTO deve ser recolhimento de IRRF !", "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQLTipoFat = Nothing


 	If VisibleMode Then

		Set Interface =CreateBennerObject("BSINTERFACE0031.Rotinas")
    	Interface.IRRFRecolhimento(CurrentSystem,CurrentQuery.FieldByName("ROTINAFIN").AsInteger)

    ElseIf WebMode Then

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		Interface.IRRFRecolhimento(CurrentSystem,CurrentQuery.FieldByName("ROTINAFIN").AsInteger, CurrentQuery.FieldByName("NOMEARQUIVO").AsString)

    End If

    Set Interface =Nothing

	Exit Sub

	erro:
		bsShowMessage(Error, "E")
		CanContinue = False
End Sub
