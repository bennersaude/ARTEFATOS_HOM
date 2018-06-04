'HASH: 9C4D0C8AE4F21BC1771A92986938041F
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If WebMode Then
		NOMEARQUIVO.Visible = True
		ROTAVISO.Visible = True
	ElseIf VisibleMode Then
		NOMEARQUIVO.Visible = False
		ROTAVISO.Visible = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim vsOpcao As String

	If CurrentQuery.FieldByName("RADIOOPCAO").AsInteger = 1 Then
		vsOpcao = "A"
	Else
		vsOpcao = "S"
	End If


	Dim Obj As Object

	On Error GoTo erro

	If VisibleMode Then
		Set Obj = CreateBennerObject("BSINTERFACE0031.Rotinas")

		Obj.IRRFaRecolher(CurrentSystem,CurrentQuery.FieldByName("DATAINICIAL").AsDateTime,CurrentQuery.FieldByName("DATAFINAL").AsDateTime,vsOpcao)

	Else
		If CurrentQuery.FieldByName("NOMEARQUIVO").IsNull Then
			bsShowMessage("Necessário informar o nome do arquivo!", "E")
			CanContinue = False
			Exit Sub
		End If
		Dim vsAux As String

		Set Obj = CreateBennerObject("BSFIN007.Rotinas")
		vsAux = Obj.IRRFaRecolher(CurrentSystem,CurrentQuery.FieldByName("DATAINICIAL").AsDateTime,CurrentQuery.FieldByName("DATAFINAL").AsDateTime,vsOpcao,CurrentQuery.FieldByName("NOMEARQUIVO").AsString, 0)
		If vsAux <> "" Then
			bsShowMessage(vsAux, "E")
			CanContinue = False
			Exit Sub
		End If
	End If


	Set Obj = Nothing

	Exit Sub

	erro:
		bsShowMessage(Error, "E")
		CanContinue = False

End Sub
