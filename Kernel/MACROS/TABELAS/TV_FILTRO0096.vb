'HASH: 40CB5B08206046FD5BA1E06E0BB5D409
 
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


	Dim vsOpcao As String

	If CurrentQuery.FieldByName("RADIOOPCAO").AsInteger = 0 Then
		vsOpcao ="G" 'Geral
	Else
		vsOpcao ="P" 'por Plano
	End If

 	If VisibleMode Then

		Set Interface =CreateBennerObject("BSINTERFACE0031.Rotinas")
    	Interface.ISSaRecolher(CurrentSystem,CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime, vsOpcao)

    ElseIf WebMode Then

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		Interface.ISSaRecolher(CurrentSystem,CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime, vsOpcao , CurrentQuery.FieldByName("NOMEARQUIVO").AsString)

    End If

    Set Interface =Nothing

	Exit Sub

	erro:
		bsShowMessage(Error, "E")
		CanContinue = False
End Sub
