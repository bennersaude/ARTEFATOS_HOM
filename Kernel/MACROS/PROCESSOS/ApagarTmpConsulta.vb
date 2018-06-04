'HASH: 72A320C44F629EC51B3C02B42D8292BC

Public Sub Main
	Dim minChave As Long
	Dim maxChave As Long
	Dim atualChave As Long
	Dim numeroApagar As Long
	Dim qPrimeiroRegistro As Object
	Dim qRemoveRegistro As Object

	Set qPrimeiroRegistro = NewQuery
	qPrimeiroRegistro.Active = False
	qPrimeiroRegistro.Add("SELECT MIN(CHAVE) CHAVE FROM TMP_CONSULTA")
	qPrimeiroRegistro.Active = True

	minChave = qPrimeiroRegistro.FieldByName("CHAVE").AsInteger
	NewCounter2("AUTORIZADORSP", 0, 1, maxChave)

	If (maxChave - minChave > 100) Then
		maxChave = maxChave - 100

		Set qRemoveRegistro = NewQuery
		qRemoveRegistro.Add("DELETE FROM TMP_CONSULTA WHERE CHAVE = :CHAVE")

		For atualChave = minChave To maxChave
			qRemoveRegistro.ParamByName("CHAVE").AsInteger = atualChave
			qRemoveRegistro.ExecSQL
		Next
	End If

End Sub
