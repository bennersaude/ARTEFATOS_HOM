'HASH: F42E685908646E6A8962C9BD1BA8915D
 
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
    	Interface.MensalidadeProRata(CurrentSystem, CurrentQuery.FieldByName("COMPETENCIA").AsDateTime, CurrentQuery.FieldByName("TIPODATA").AsString)

    ElseIf WebMode Then
    	Dim vsAux As String

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		vsAux = Interface.MensalidadeProRata(CurrentSystem, _
											 CurrentQuery.FieldByName("COMPETENCIA").AsDateTime, _
											 CurrentQuery.FieldByName("TIPODATA").AsString, _
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
