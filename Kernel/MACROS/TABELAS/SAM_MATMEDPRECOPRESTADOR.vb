'HASH: CDD7167997D411FB59066C272B9C35EE
'#Uses "*bsShowMessage"

'início sms 62791 - Edilson.Castro - 01/09/2006

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebMenuCode = "T4245" Then
			MATMEDPRECOTAB.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim q As Object
	Set q = NewQuery

	q.Add("SELECT 1")
	q.Add("  FROM SAM_MATMEDPRECOESTADO")
	q.Add(" WHERE MATMEDPRECOTAB = :HandleTabela")

	q.ParamByName("HandleTabela").AsInteger = CurrentQuery.FieldByName("MATMEDPRECOTAB").AsInteger
	q.Active = True

	CanContinue = q.EOF

	q.Active = False

	Set q = Nothing

	If Not CanContinue Then
		bsShowMessage("Esta tabela já está vinculada a estados, não é possível cadastrar prestadores.", "I")
		Exit Sub
	End If

	Dim Interface As Object
	Dim Linha As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_MATMEDPRECOPRESTADOR", "DATAINICIAL", "DATAFINAL", CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime, "PRESTADOR", "")

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
		Exit Sub
	End If
End Sub
'fim sms 62791
