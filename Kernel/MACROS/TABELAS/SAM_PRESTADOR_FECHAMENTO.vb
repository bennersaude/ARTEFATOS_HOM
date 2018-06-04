'HASH: FB0CD17AE0BD33F29E074F39284C92EB
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If WebMode Then
		If WebVisionCode = "W_SAM_PRESTADOR_FECHAMENTO" Then
			Dim VSQL As Object

			If CurrentQuery.FieldByName("PRIMEIRODIA").AsInteger = 0 And _
				CurrentQuery.FieldByName("SEGUNDODIA").AsInteger = 0 And _
				CurrentQuery.FieldByName("TERCEIRODIA").AsInteger = 0 And _
				CurrentQuery.FieldByName("QUARTODIA").AsInteger = 0 And _
				CurrentQuery.FieldByName("QUINTODIA").AsInteger = 0 Then
				CancelDescription = "Por favor, preencha pelo menos um dos campos."
				CanContinue = False
			Else
				Set VSQL = NewQuery

				VSQL.Clear

				VSQL.Add("SELECT 1 QTD FROM SAM_PRESTADOR_FECHAMENTO WHERE PRESTADOR = :HPREST")

				VSQL.ParamByName("HPREST").AsInteger = CInt(SessionVar("HPRESTADOR"))
				VSQL.Active = True

				If (VSQL.FieldByName("QTD").AsInteger = 1) And (CurrentQuery.State = 3) Then
					CancelDescription = "Já existe registro para esse prestador. Não é possível continuar."
					CanContinue = False
				Else
					CurrentQuery.FieldByName("PRESTADOR").AsInteger = CInt(SessionVar("HPRESTADOR"))
				End If

				Set VSQL = Nothing
			End If
		End If

		Exit Sub
	End If

	If (CurrentQuery.FieldByName("PRIMEIRODIA").IsNull) And _
	   (Not CurrentQuery.FieldByName("SEGUNDODIA").IsNull Or _
		Not CurrentQuery.FieldByName("TERCEIRODIA").IsNull Or _
		Not CurrentQuery.FieldByName("QUARTODIA").IsNull Or _
		Not CurrentQuery.FieldByName("QUINTODIA").IsNull) Then
		bsShowMessage("O primeiro dia deve ser preenchido antes de qualquer outro.", "E")
		CurrentQuery.FieldByName("SEGUNDODIA").Value = Null
		CurrentQuery.FieldByName("TERCEIRODIA").Value = Null
		CurrentQuery.FieldByName("QUARTODIA").Value = Null
		CurrentQuery.FieldByName("QUINTODIA").Value = Null
		PRIMEIRODIA.SetFocus
		CanContinue = False
		Exit Sub
	End If

	If (CurrentQuery.FieldByName("SEGUNDODIA").IsNull) And _
	   (Not CurrentQuery.FieldByName("TERCEIRODIA").IsNull Or _
		Not CurrentQuery.FieldByName("QUARTODIA").IsNull Or _
		Not CurrentQuery.FieldByName("QUINTODIA").IsNull) Then
		bsShowMessage("O segundo dia deve ser preenchido antes.", "E")
		CurrentQuery.FieldByName("TERCEIRODIA").Value = Null
		CurrentQuery.FieldByName("QUARTODIA").Value = Null
		CurrentQuery.FieldByName("QUINTODIA").Value = Null
		SEGUNDODIA.SetFocus
		CanContinue = False
		Exit Sub
	End If

	If (CurrentQuery.FieldByName("TERCEIRODIA").IsNull) And _
	   (Not CurrentQuery.FieldByName("QUARTODIA").IsNull Or _
		Not CurrentQuery.FieldByName("QUINTODIA").IsNull) Then
		bsShowMessage("O terceiro dia deve ser preenchido antes.", "E")
		CurrentQuery.FieldByName("QUARTODIA").Value = Null
		CurrentQuery.FieldByName("QUINTODIA").Value = Null
		TERCEIRODIA.SetFocus
		CanContinue = False
		Exit Sub
	End If

	If (CurrentQuery.FieldByName("QUARTODIA").IsNull) And (Not CurrentQuery.FieldByName("QUINTODIA").IsNull) Then
		bsShowMessage("O terceiro dia deve ser preenchido antes.", "E")
		CurrentQuery.FieldByName("QUINTODIA").Value = Null
		QUARTODIA.SetFocus
		CanContinue = False
		Exit Sub
	End If

	'Primeiro dia
	If Not CurrentQuery.FieldByName("SEGUNDODIA").IsNull Then
		If (CurrentQuery.FieldByName("PRIMEIRODIA").AsInteger >= CurrentQuery.FieldByName("SEGUNDODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("SEGUNDODIA").Value = Null
			SEGUNDODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("TERCEIRODIA").IsNull Then
		If (CurrentQuery.FieldByName("PRIMEIRODIA").AsInteger >= CurrentQuery.FieldByName("TERCEIRODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("TERCEIRODIA").Value = Null
			TERCEIRODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("QUARTODIA").IsNull Then
		If (CurrentQuery.FieldByName("PRIMEIRODIA").AsInteger >= CurrentQuery.FieldByName("QUARTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUARTODIA").Value = Null
			QUARTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("QUINTODIA").IsNull Then
		If (CurrentQuery.FieldByName("PRIMEIRODIA").AsInteger >= CurrentQuery.FieldByName("QUINTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUINTODIA").Value = Null
			QUINTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	'Segundo dia
	If Not CurrentQuery.FieldByName("TERCEIRODIA").IsNull Then
		If (CurrentQuery.FieldByName("SEGUNDODIA").AsInteger >= CurrentQuery.FieldByName("TERCEIRODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("TERCEIRODIA").Value = Null
			TERCEIRODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("QUARTODIA").IsNull Then
		If (CurrentQuery.FieldByName("SEGUNDODIA").AsInteger >= CurrentQuery.FieldByName("QUARTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUARTODIA").Value = Null
			QUARTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("QUINTODIA").IsNull Then
		If (CurrentQuery.FieldByName("SEGUNDODIA").AsInteger >= CurrentQuery.FieldByName("QUINTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUINTODIA").Value = Null
			QUINTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	'Terceiro dia
	If Not CurrentQuery.FieldByName("QUARTODIA").IsNull Then
		If (CurrentQuery.FieldByName("TERCEIRODIA").AsInteger >= CurrentQuery.FieldByName("QUARTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUARTODIA").Value = Null
			QUARTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("QUINTODIA").IsNull Then
		If (CurrentQuery.FieldByName("TERCEIRODIA").AsInteger >= CurrentQuery.FieldByName("QUINTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUINTODIA").Value = Null
			QUINTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If

	'Quarto dia
	If Not CurrentQuery.FieldByName("QUINTODIA").IsNull Then
		If (CurrentQuery.FieldByName("QUARTODIA").AsInteger >= CurrentQuery.FieldByName("QUINTODIA").AsInteger) Then
			bsShowMessage("Não pode existir dia menor ou igual ao anterior.", "E")
			CurrentQuery.FieldByName("QUINTODIA").Value = Null
			QUINTODIA.SetFocus
			CanContinue = False
			Exit Sub
		End If
	End If
End Sub
