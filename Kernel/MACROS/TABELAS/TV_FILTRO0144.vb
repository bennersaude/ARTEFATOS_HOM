'HASH: 70D17B15799A324F005F377F8B605963

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

	On Error GoTo erro

	If VisibleMode Then

		Set Interface = CreateBennerObject("BSINTERFACE0031.Rotinas")
    	Interface.DefasagemPagtoPlano(CurrentSystem, _
    								  CurrentQuery.FieldByName("COMPETENCIA").AsDateTime, _
    								  CurrentQuery.FieldByName("PLANO").AsInteger, _
    								  CurrentQuery.FieldByName("FILIAL").AsInteger)

    ElseIf WebMode Then
    	Dim vsAux As String

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		vsAux = Interface.DefasagemPagtoPlano(CurrentSystem, _
											  CurrentQuery.FieldByName("COMPETENCIA").AsDateTime, _
											  CurrentQuery.FieldByName("PLANO").AsInteger, _
									 		  CurrentQuery.FieldByName("FILIAL").AsInteger, _
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
