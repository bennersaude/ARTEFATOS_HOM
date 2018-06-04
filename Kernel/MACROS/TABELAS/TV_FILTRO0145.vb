'HASH: C43EC5AE65192D72ACDC0692F7A13842

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
    	Interface.ComposicaoSaldo(CurrentSystem, _
								  CurrentQuery.FieldByName("OPCAO").AsString, _
							 	  CurrentQuery.FieldByName("TIPODATA").AsString, _
							 	  CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
							 	  CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
							 	  CurrentQuery.FieldByName("PERCENTUALCPMF").AsFloat, _
							 	  CurrentQuery.FieldByName("TARIFALANC").AsCurrency, _
							 	  CurrentQuery.FieldByName("TESOURARIA").AsInteger)

    ElseIf WebMode Then
    	Dim vsAux As String

    	If CurrentQuery.FieldByName("NOMEARQUIVO").AsString = "" Then
    		bsShowMessage("Nome do Arquivo obrigatório!", "E")
    		CanContinue = False
    		Exit Sub
    	End If

		Set Interface = CreateBennerObject("BSFIN007.Rotinas")
		vsAux = Interface.ComposicaoSaldo(CurrentSystem, _
										  CurrentQuery.FieldByName("OPCAO").AsString, _
									 	  CurrentQuery.FieldByName("TIPODATA").AsString, _
									 	  CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
									 	  CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
									 	  CurrentQuery.FieldByName("PERCENTUALCPMF").AsFloat, _
									 	  CurrentQuery.FieldByName("TARIFALANC").AsCurrency, _
									 	  CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
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

