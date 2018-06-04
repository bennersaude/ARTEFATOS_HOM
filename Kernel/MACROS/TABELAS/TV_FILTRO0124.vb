'HASH: 3415CB9E26E5713A965BC04D700F9E9B
 '#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull) And (Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull) Then
		If (CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime) > (CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime) Then
			bsShowMessage("Competência Inicial é maior que a  Competência Final", "E")
			CanContinue = False
			Exit Sub
		End If

	If CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime = CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime Then
		SessionVar("vTextoFiltro") = "Competência : " + Str(Format(CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime,"mm/yyyy"))
	Else
		vTextoFiltro = "Competência entre " + Str(Format(CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime,"mm/yyyy")) + " e " + _
		Str(Format(CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime,"mm/yyyy"))
	End If

	ElseIf (Not CurrentQuery.FieldByName("COMPETENCIAINICIAL").IsNull) Then
		SessionVar("vTextoFiltro") = "Competência a partir de " + Str(Format(CurrentQuery.FieldByName("COMPETENCIAINICIAL").AsDateTime,"mm/yyyy"))
	ElseIf (Not CurrentQuery.FieldByName("COMPETENCIAFINAL").IsNull) Then
		SessionVar("vTextoFiltro") = "Competência até " + Str(Format(CurrentQuery.FieldByName("COMPETENCIAFINAL").AsDateTime,"mm/yyyy"))
	End If
End Sub
