'HASH: EFC1BDFD16FBD72109CF97A209C2133A
'Macro: SAM_PRESTADOR_ISS
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()
	TABLE_AfterScroll
End Sub

Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
		COMPETENCIAFINAL.ReadOnly = False
	Else
		COMPETENCIAFINAL.ReadOnly = True
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Interface As Object
	Dim Linha As String
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_ISS", "COMPETENCIAINICIAL", "COMPETENCIAFINAL", CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime, CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime, "PRESTADOR", "")

	If Linha <> "" Then
		CanContinue = False
		bsShowMessage(Linha, "E")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("TABRECOLHIMENTO").AsInteger = 1 Then
		CurrentQuery.FieldByName("MOTIVOISENCAO").Clear

		If (Not CurrentQuery.FieldByName("ALIQUOTA").IsNull) And _
		   (CurrentQuery.FieldByName("ALIQUOTA").AsFloat = 0) Then
			CanContinue = False
			bsShowMessage("Percentual = 0 deve ser considerado como isento. Se for para considerar o percentual da cidade, deixe sem nenhum valor em percentual", "E")
			Exit Sub
		End If
	Else
		CurrentQuery.FieldByName("ALIQUOTA").Clear
		CurrentQuery.FieldByName("ABATEMATERIAL").Value = "N"
		CurrentQuery.FieldByName("ABATEMEDICAMENTO").Value = "N"
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	If Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull Then
		CanContinue = False
		bsShowMessage("Registro finalizado não pode ser alterado!", "E")
		Exit Sub
	End If
End Sub
