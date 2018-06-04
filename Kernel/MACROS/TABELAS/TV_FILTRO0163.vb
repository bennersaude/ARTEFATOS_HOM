'HASH: 3DEB9F619D4C5D2C3291664B3DAA08C0
Option Explicit

Public Function ProcuraBeneficiario(pEntrada As String, pHandleOuCPF As Boolean)
	Dim qBeneficiario As BPesquisa
	Set qBeneficiario = NewQuery

	qBeneficiario.Active = False
	qBeneficiario.Add(" SELECT B.HANDLE BENEFICIARIO, M.CPF            ")
	qBeneficiario.Add(" FROM SAM_BENEFICIARIO B                        ")
	qBeneficiario.Add(" JOIN SAM_MATRICULA M ON M.HANDLE = B.MATRICULA ")

	If pHandleOuCPF Then
		qBeneficiario.Add(" WHERE B.HANDLE = :HANDLE                   ")
		qBeneficiario.ParamByName("HANDLE").AsString = pEntrada
		qBeneficiario.Active = True

		CurrentQuery.FieldByName("CPF").AsString = ""

		If (qBeneficiario.FieldByName("CPF").AsString <> "") Then
			CurrentQuery.FieldByName("CPF").AsString = qBeneficiario.FieldByName("CPF").AsString
		End If

	Else
		qBeneficiario.Add(" WHERE M.CPF = :CPF                         ")
		qBeneficiario.ParamByName("CPF").AsString = pEntrada
		qBeneficiario.Active = True

		CurrentQuery.FieldByName("BENEFICIARIO").AsString = ""

		If (qBeneficiario.FieldByName("BENEFICIARIO").AsInteger > 0) Then
			CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = qBeneficiario.FieldByName("BENEFICIARIO").AsInteger
		End If
	End If

	Set qBeneficiario = Nothing

End Function

Public Function ProcuraPessoa(pEntrada As String, pHandleOuCPF As Boolean)
	Dim qPessoa As BPesquisa
	Set qPessoa = NewQuery

	qPessoa.Active = False
	qPessoa.Add(" SELECT P.HANDLE PESSOA, P.CNPJCPF CPF  ")
	qPessoa.Add(" FROM SFN_PESSOA P                      ")

	If pHandleOuCPF Then
		qPessoa.Add(" WHERE P.HANDLE = :HANDLE           ")
		qPessoa.ParamByName("HANDLE").AsString = pEntrada
		qPessoa.Active = True

		CurrentQuery.FieldByName("CPF").AsString = ""

		If (qPessoa.FieldByName("CPF").AsString <> "") Then
			CurrentQuery.FieldByName("CPF").AsString = qPessoa.FieldByName("CPF").AsString
		End If
	Else
		qPessoa.Add(" WHERE P.CNPJCPF = :CPF             ")
		qPessoa.ParamByName("CPF").AsString = pEntrada
		qPessoa.Active = True

		CurrentQuery.FieldByName("PESSOA").AsString = ""

 		If (qPessoa.FieldByName("PESSOA").AsInteger > 0) Then
			CurrentQuery.FieldByName("PESSOA").AsInteger = qPessoa.FieldByName("PESSOA").AsInteger
		End If
	End If

	Set qPessoa = Nothing
End Function

Public Sub BENEFICIARIO_OnExit()
	Dim HBeneficiario As String
	Dim CPFFormatado As String

	HBeneficiario = CurrentQuery.FieldByName("BENEFICIARIO").AsString

	If HBeneficiario = "" Then
		Exit Sub
	End If

	ProcuraBeneficiario(HBeneficiario, True)
	ProcuraPessoa(CurrentQuery.FieldByName("CPF").AsString, False)

	CPFFormatado = CurrentQuery.FieldByName("CPF").AsString
	CPFFormatado = Mid(CPFFormatado, 1, 3) + "." + Mid(CPFFormatado, 4, 3) + "." + Mid(CPFFormatado, 7, 3) + "-" + Mid(CPFFormatado, 10, 2)

	CurrentQuery.FieldByName("CPF").AsString = CPFFormatado
End Sub

Public Sub CPF_OnExit()

	Dim CpfInformado As String

	CpfInformado = Replace(CurrentQuery.FieldByName("CPF").AsString, "-", "")
	CpfInformado = Replace(CpfInformado, ".", "")

	If CpfInformado = "" Then
		Exit Sub
	End If

	ProcuraPessoa(CpfInformado, False)
	ProcuraBeneficiario(CpfInformado, False)

End Sub

Public Sub PESSOA_OnExit()
	Dim HPessoa As String
	Dim CPFFormatado As String

	HPessoa = CurrentQuery.FieldByName("PESSOA").AsString

	If HPessoa = "" Then
		Exit Sub
	End If

	ProcuraPessoa(HPessoa, True)

	ProcuraBeneficiario(CurrentQuery.FieldByName("CPF").AsString, False)

	CPFFormatado = CurrentQuery.FieldByName("CPF").AsString
	CPFFormatado = Mid(CPFFormatado, 1, 3) + "." + Mid(CPFFormatado, 4, 3) + "." + Mid(CPFFormatado, 7, 3) + "-" + Mid(CPFFormatado, 10, 2)

	CurrentQuery.FieldByName("CPF").AsString = CPFFormatado
End Sub

Public Sub TABLE_AfterScroll()
	Dim CPFFormatado As String

	If SessionVar("DMEDCPF") <> "" Then
		CPFFormatado = SessionVar("DMEDCPF")
		CPFFormatado = Mid(CPFFormatado, 1, 3) + "." + Mid(CPFFormatado, 4, 3) + "." + Mid(CPFFormatado, 7, 3) + "-" + Mid(CPFFormatado, 10, 2)
		CurrentQuery.FieldByName("CPF").AsString = CPFFormatado

		CurrentQuery.FieldByName("ANO").AsInteger = CInt(SessionVar("DMEDANO"))

		CPF_OnExit

		SessionVar("DMEDCPF") = ""
		SessionVar("DMEDANO") = ""
	End If
End Sub

