'HASH: 785B380CD35207B593DD73CDCFC99FAE
''''#uses "*FormatarTelefone"

'SMS 78298 - Paulo Drummond Filho - 05/04/2007
' A macro foi modificada em vários lugares para corrigir problemas de máscara.
' Comentar os lugares onde houve modificação causaria uma poluição visual muito
' grande tornando assim muito confusa a leitura da macro.
Public Sub TABLE_AfterEdit()
	CurrentQuery.FieldByName("DATAATUALIZACAO").AsDateTime = ServerNow
	If WebMode Then
		CurrentQuery.FieldByName("TELEFONE1").Mask = ""
		CurrentQuery.FieldByName("FAX").Mask = ""
		CurrentQuery.FieldByName("CELULAR").Mask = ""
	End If
End Sub

Public Sub TABLE_AfterInsert()
	If WebMode Then
		CurrentQuery.FieldByName("TELEFONE1").Mask = ""
		CurrentQuery.FieldByName("FAX").Mask = ""
		CurrentQuery.FieldByName("CELULAR").Mask = ""
	End If
End Sub

Public Sub TABLE_AfterPost()
	InfoDescription = "A alteração foi confirmada. Operação concluída"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
 ' Mochi TJDF

	If WebMode And WebMenuCode <> "" Then
		Dim vDDD As String
		Dim vTelefoneNovo As String
		Dim vTelefone As String

		If CurrentQuery.FieldByName("TELEFONE1").AsString <> "" Then
			If InStr(CurrentQuery.FieldByName("TELEFONE1").AsString,")") > 0 Then
				vDDD = Mid(CurrentQuery.FieldByName("TELEFONE1").AsString,2,InStr(CurrentQuery.FieldByName("TELEFONE1").AsString,")")-2)
				vDDD = Trim(vDDD)
			Else
				vDDD = ""
			End If

			vTelefone = Mid(CurrentQuery.FieldByName("TELEFONE1").AsString,InStr(CurrentQuery.FieldByName("TELEFONE1").AsString,")")+1, Len(CurrentQuery.FieldByName("TELEFONE1").AsString)-InStr(CurrentQuery.FieldByName("TELEFONE1").AsString,")"))
			vTelefone = Trim(vTelefone)

			On Error GoTo TrataErroTelefone
			vTelefoneNovo = FormatarTelefone(vDDD,vTelefone)
			If InStr(vTelefoneNovo,"(") = 1 Then
				CurrentQuery.FieldByName("TELEFONE1").AsString = vTelefoneNovo
			Else
				CanContinue = False
				CancelDescription = "No campo 'Telefone' o " + vTelefoneNovo
				Exit Sub
			End If

			GoTo TrataFAX
			TrataErroTelefone:
				CanContinue = False
				CancelDescription = "Ocorreu um erro durante o tratamento do campo 'Telefone'. Por favor, verifique o campo."
				Exit Sub
		End If

		TrataFAX:
		'vALIDANDO FAX
		If CurrentQuery.FieldByName("FAX").AsString <> "" Then
			vDDD = ""
			vTelefoneNovo = ""
			If InStr(CurrentQuery.FieldByName("FAX").AsString,")") > 0 Then
				vDDD = Mid(CurrentQuery.FieldByName("FAX").AsString,2,InStr(CurrentQuery.FieldByName("FAX").AsString,")")-2)
				vDDD = Trim(vDDD)
			Else
				vDDD = ""
			End If

			vTelefone = Mid(CurrentQuery.FieldByName("FAX").AsString,InStr(CurrentQuery.FieldByName("FAX").AsString,")")+1, Len(CurrentQuery.FieldByName("FAX").AsString)-InStr(CurrentQuery.FieldByName("FAX").AsString,")"))
			vTelefone = Trim(vTelefone)

			On Error GoTo TrataErroFAX
			vTelefoneNovo = FormatarTelefone(vDDD,vTelefone)
			If InStr(vTelefoneNovo,"(") = 1 Then
				CurrentQuery.FieldByName("FAX").AsString = vTelefoneNovo
			Else
				CanContinue = False
				CancelDescription = "No campo 'FAX' o " + vTelefoneNovo
				Exit Sub
			End If

			GoTo TrataCelular
			TrataErroFAX:
				CanContinue = False
				CancelDescription = "Ocorreu um erro durante o tratamento do campo 'FAX'. Por favor, verifique o campo."
				Exit Sub
		End If

		TrataCelular:
		'Validando Celular
		If CurrentQuery.FieldByName("CELULAR").AsString <> "" Then
			vDDD = ""
			vTelefoneNovo = ""
			If InStr(CurrentQuery.FieldByName("CELULAR").AsString,")") > 0 Then
				vDDD = Mid(CurrentQuery.FieldByName("CELULAR").AsString,2,InStr(CurrentQuery.FieldByName("CELULAR").AsString,")")-2)
				vDDD = Trim(vDDD)
			Else
				vDDD = ""
			End If

			vTelefone = Mid(CurrentQuery.FieldByName("CELULAR").AsString, InStr(CurrentQuery.FieldByName("CELULAR").AsString,")")+1, Len(CurrentQuery.FieldByName("CELULAR").AsString)-InStr(CurrentQuery.FieldByName("CELULAR").AsString,")"))
			vTelefone = Trim(vTelefone)

			vTelefoneNovo = FormatarTelefone(vDDD,vTelefone)

			On Error GoTo TrataErroCelular
			If InStr(vTelefoneNovo,"(") = 1 Then
				CurrentQuery.FieldByName("CELULAR").AsString = vTelefoneNovo
			Else
				CanContinue = False
				CancelDescription = "No campo 'Celular' o " + vTelefoneNovo
				Exit Sub
			End If

			Exit Sub
			TrataErroCelular:
				CanContinue = False
				CancelDescription = "Ocorreu um erro durante o tratamento do campo 'Celular'. Por favor, verifique o campo."
				Exit Sub
		End If
	End If
End Sub

Public Function FormatarTelefone(DDD As String, Telefone As String) As String
	Dim Aux As String
	Dim AuxTel As String
	Dim Prefixo As String
	Dim Sufixo As String
	Dim i As Integer
	Dim j As Integer

	AuxTel = "(" + DDD + ")"

	j = 1
	i = 1

	If DDD = "" Then
		FormatarTelefone = "telefone não é válido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
		Exit Function
	Else
		'A máscara dos campos de telefone permite que o DDD possua 4 digitos
		If Len(DDD) > 4 Then
			FormatarTelefone = "telefone não é válido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
			Exit Function
		Else
			While (i <= Len(DDD))
				Aux =  Mid(Telefone,i,1)
				If Not ((Aux = "1") Or (Aux = "2") Or _
				        (Aux = "3") Or (Aux = "4") Or _
				        (Aux = "5") Or (Aux = "6") Or _
				        (Aux = "7") Or (Aux = "8") Or _
				        (Aux = "9") Or (Aux = "0")) Then
					FormatarTelefone = "telefone não é valido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
					Exit Function
				End If
				i = i + 1
			Wend
		End If
	End If

	j = 1
	i = 1
	While (i <= Len(Telefone)) And (Mid(Telefone,i,1) <> "-")
		Aux =  Mid(Telefone,i,1)
		If (Aux = "1") Or _
		(Aux = "2") Or _
		(Aux = "3") Or _
		(Aux = "4") Or _
		(Aux = "5") Or _
		(Aux = "6") Or _
		(Aux = "7") Or _
		(Aux = "8") Or _
		(Aux = "9") Or _
		(Aux = "0") Then
			If j < 5 Then
				Prefixo = Prefixo + Aux
			Else
				If j = 5 Then
					Prefixo = Prefixo + "-"
					'Prefixo = Prefixo + Aux
					Sufixo = Sufixo + Aux
				Else
					'Prefixo = Prefixo + Aux
					Sufixo = Sufixo + Aux
				End If
			End If
		Else
			' Encontrou algum caractere estranho no telefone
			FormatarTelefone = "telefone não é válido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
			Exit Function
		End If
		i = i + 1
		j = j + 1
	Wend


	If i < Len(Telefone) Then 'significa que o usuário digitou "-" no telefone.
		' Verifica se o prefixo tem 4 dígitos
		If Len(Prefixo) <> 4 Then
			FormatarTelefone = "telefone não é válido. Prefixo deve ter 4 dígitos"
			Exit Function
		End If
		Aux = Mid(Telefone,i+1,Len(Telefone)-i)
		If Len(Aux) <> 4 Then
			FormatarTelefone = "telefone não é válido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
			Exit Function
		End If
		Dim aux2 As String
		j = 1
		While j < Len(Aux)
		aux2 = Mid(Aux,j,1)
			If (aux2 <> "1") And _
			(aux2 <> "2") And _
			(aux2 <> "3") And _
			(aux2 <> "4") And _
			(aux2 <> "5") And _
			(aux2 <> "6") And _
			(aux2 <> "7") And _
			(aux2 <> "8") And _
			(aux2 <> "9") And _
			(aux2 <> "0") Then
				FormatarTelefone = "telefone não é válido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
				Exit Function
			End If
			j = j + 1
		Wend
		AuxTel = AuxTel + Prefixo + "-" + Aux
	Else
		If Len(Sufixo) <> 4 Then
			FormatarTelefone = "telefone não é válido. Por favor digite o número no seguinte modelo: '(99)9999-9999'"
			Exit Function
		End If
		AuxTel = AuxTel + Prefixo + Sufixo
	End If

	FormatarTelefone = AuxTel

End Function
