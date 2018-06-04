'HASH: B5AEEEBDD3009B83CBA57D10C5977F71

'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If VisibleMode Then
		NOMEARQUIVO.Visible = False
		ROTAVISO.Visible = False
	ElseIf WebMode Then
		NOMEARQUIVO.Required = True
		NOMEARQUIVO.Visible = True
		ROTAVISO.Visible = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object

	If CurrentQuery.FieldByName("VALORINICIAL").AsFloat <= 0 Then
		bsShowMessage("O valor deve ser maior que 0", "E")
		CanContinue = False
		Exit Sub
	End If

	On Error GoTo erro

	If VisibleMode Then

		Set Interface =CreateBennerObject("BSINTERFACE0031.Rotinas")
    	Interface.RateioCPMF(CurrentSystem, _
    						CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
    						CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
    						CurrentQuery.FieldByName("VALORINICIAL").AsFloat)

    ElseIf WebMode Then
    	Dim vsAux As String

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		vsAux = Interface.RateioCPMF(CurrentSystem, _
					  				 CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
									 CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
							 		 CurrentQuery.FieldByName("VALORINICIAL").AsFloat, _
							 		 CurrentQuery.FieldByName("NOMEARQUIVO").AsString)

		If vsAux <> "" Then
			bsShowMessage(vsAux, "E")
			CanContinue = False
			Exit Sub
		End If
    End If

    Set Interface = Nothing

	Exit Sub

	erro:
		bsShowMessage(Error, "E")
		CanContinue = False
End Sub
