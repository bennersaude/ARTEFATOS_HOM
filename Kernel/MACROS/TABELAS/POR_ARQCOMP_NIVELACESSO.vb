'HASH: F063AF0DC318B3AFB8AFEA4E136FFF62
 
Public Sub TABLE_AfterScroll()
	If Not(CurrentQuery.FieldByName("TABNIVEL").AsInteger = 1) Then

		If Not(CurrentQuery.FieldByName("USUARIO").IsNull) Then
			Dim query As BPesquisa

			Set query = NewQuery

			query.Clear
			query.Add("SELECT P.TIPO, U.CODIGO FROM POR_PERFILUSUARIO P ")
			query.Add(" JOIN POR_GRUPOSEGURANCA S ON S.PERFILUSUARIO = P.HANDLE ")
			query.Add(" JOIN POR_USUARIO U ON U.GRUPOSEGURANCA = S.HANDLE ")
			query.Add("  WHERE U.HANDLE = :USUARIO ")
			query.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("USUARIO").Value
			query.Active = True


			If Not (query.EOF) Then

				Dim nome As String
				Dim codigo As String

				codigo = query.FieldByName("CODIGO").AsString

				If (query.FieldByName("TIPO").AsInteger = 1) Then

					query.Clear
					query.Add("SELECT NOME FROM SAM_BENEFICIARIO ")
					query.Add("WHERE HANDLE = :HANDLE")
					query.ParamByName("HANDLE").AsString = codigo
					query.Active = True

					If Not (query.EOF) Then
						nome = query.FieldByName("NOME").AsString
					End If

				ElseIf (query.FieldByName("TIPO").AsInteger = 2) Then

					query.Clear
					query.Add("SELECT NOME FROM SAM_PRESTADOR ")
					query.Add("WHERE HANDLE = :HANDLE")
					query.ParamByName("HANDLE").AsString = codigo
					query.Active = True

					If Not (query.EOF) Then
						nome = query.FieldByName("NOME").AsString
					End If


				ElseIf (query.FieldByName("TIPO").AsInteger = 3) Then

					query.Clear
					query.Add("SELECT C.* FROM SAM_CONTRATO C ")
					query.Add(" WHERE C.HANDLE IN (SELECT F.CONTRATO ")
					query.Add("  FROM SAM_FAMILIA F ")
					query.Add("   WHERE F.TITULARRESPONSAVEL = :RESPONSAVEL) ")
					query.Add("  AND (C.DATACANCELAMENTO IS NULL Or C.DATACANCELAMENTO > :DATAATUAL) ")
					query.ParamByName("RESPONSAVEL").AsString = codigo
					query.ParamByName("DATAATUAL").AsDateTime = ServerNow
					query.Active = True

					If Not (query.EOF) Then

						query.Clear
						query.Add("SELECT NOME FROM SAM_BENEFICIARIO ")
						query.Add("WHERE HANDLE = :HANDLE")
						query.ParamByName("HANDLE").AsString = codigo
						query.Active = True

						If Not (query.EOF) Then
							nome = query.FieldByName("NOME").AsString
						End If

					Else

						query.Clear
						query.Add("SELECT NOME FROM SFN_PESSOA ")
						query.Add("WHERE HANDLE = :HANDLE")
						query.ParamByName("HANDLE").AsString = codigo
						query.Active = True

						If Not (query.EOF) Then
							nome = query.FieldByName("NOME").AsString
						End If

					End If


				ElseIf (query.FieldByName("TIPO").AsInteger = 5) Then

					query.Clear
					query.Add("SELECT NOME FROM Z_GRUPOUSUARIOS ")
					query.Add("WHERE HANDLE = :HANDLE")
					query.ParamByName("HANDLE").AsString = codigo
					query.Active = True

					If Not (query.EOF) Then
						nome = query.FieldByName("NOME").AsString
					End If


				End If

				NOMEUSUARIO.Text = nome

			End If

			Set query = Nothing

		Else

			NOMEUSUARIO.Text = ""

		End If

	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

	If (CurrentQuery.FieldByName("TABNIVEL").AsInteger = 1) Then
		CurrentQuery.FieldByName("USUARIO").Clear
	Else
		CurrentQuery.FieldByName("PERFILUSUARIO").Clear
	End If

End Sub

Public Sub USUARIO_OnChange()

	If Not(CurrentQuery.FieldByName("USUARIO").IsNull) Then
		Dim query As BPesquisa

		Set query = NewQuery

		query.Clear
		query.Add("SELECT P.TIPO, U.CODIGO FROM POR_PERFILUSUARIO P ")
		query.Add(" JOIN POR_GRUPOSEGURANCA S ON S.PERFILUSUARIO = P.HANDLE ")
		query.Add(" JOIN POR_USUARIO U ON U.GRUPOSEGURANCA = S.HANDLE ")
		query.Add("  WHERE U.HANDLE = :USUARIO ")
		query.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("USUARIO").Value
		query.Active = True


		If Not (query.EOF) Then

			Dim nome As String
			Dim codigo As String

			codigo = query.FieldByName("CODIGO").AsString

			If (query.FieldByName("TIPO").AsInteger = 1) Then

				query.Clear
				query.Add("SELECT NOME FROM SAM_BENEFICIARIO ")
				query.Add("WHERE HANDLE = :HANDLE")
				query.ParamByName("HANDLE").AsString = codigo
				query.Active = True

				If Not (query.EOF) Then
					nome = query.FieldByName("NOME").AsString
				End If

			ElseIf (query.FieldByName("TIPO").AsInteger = 2) Then

				query.Clear
				query.Add("SELECT NOME FROM SAM_PRESTADOR ")
				query.Add("WHERE HANDLE = :HANDLE")
				query.ParamByName("HANDLE").AsString = codigo
				query.Active = True

				If Not (query.EOF) Then
					nome = query.FieldByName("NOME").AsString
				End If


			ElseIf (query.FieldByName("TIPO").AsInteger = 3) Then

				query.Clear
				query.Add("SELECT C.* FROM SAM_CONTRATO C ")
				query.Add(" WHERE C.HANDLE IN (SELECT F.CONTRATO ")
				query.Add("  FROM SAM_FAMILIA F ")
				query.Add("   WHERE F.TITULARRESPONSAVEL = :RESPONSAVEL) ")
				query.Add("  AND (C.DATACANCELAMENTO IS NULL Or C.DATACANCELAMENTO > :DATAATUAL) ")
				query.ParamByName("RESPONSAVEL").AsString = codigo
				query.ParamByName("DATAATUAL").AsDateTime = ServerNow
				query.Active = True

				If Not (query.EOF) Then

					query.Clear
					query.Add("SELECT NOME FROM SAM_BENEFICIARIO ")
					query.Add("WHERE HANDLE = :HANDLE")
					query.ParamByName("HANDLE").AsString = codigo
					query.Active = True

					If Not (query.EOF) Then
						nome = query.FieldByName("NOME").AsString
					End If

				Else

					query.Clear
					query.Add("SELECT NOME FROM SFN_PESSOA ")
					query.Add("WHERE HANDLE = :HANDLE")
					query.ParamByName("HANDLE").AsString = codigo
					query.Active = True

					If Not (query.EOF) Then
						nome = query.FieldByName("NOME").AsString
					End If

				End If


			ElseIf (query.FieldByName("TIPO").AsInteger = 5) Then

				query.Clear
				query.Add("SELECT NOME FROM Z_GRUPOUSUARIOS ")
				query.Add("WHERE HANDLE = :HANDLE")
				query.ParamByName("HANDLE").AsString = codigo
				query.Active = True

				If Not (query.EOF) Then
					nome = query.FieldByName("NOME").AsString
				End If


			End If

			NOMEUSUARIO.Text = nome

		End If

		Set query = Nothing

	Else

		NOMEUSUARIO.Text = ""

	End If

End Sub
