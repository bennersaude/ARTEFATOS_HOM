'HASH: 459087AED7134923C705B10C70754B43

Public Sub Main
    Dim mensagem As String
    Dim AtualizaBenef As Object

	Dim qEnderecoBeneficiario As BPesquisa
	Set qEnderecoBeneficiario = NewQuery

	qEnderecoBeneficiario.Add("SELECT ENDERECORESIDENCIAL ")
	qEnderecoBeneficiario.Add("  FROM SAM_BENEFICIARIO    ")
	qEnderecoBeneficiario.Add(" WHERE HANDLE = :HANDLE    ")
	qEnderecoBeneficiario.ParamByName("HANDLE").AsInteger = CLng(ServiceVar("HANDLEBENEFICIARIO"))
	qEnderecoBeneficiario.Active = True

    Set AtualizaBenef = CreateBennerObject("SamBeneficiario.Atualiza")
    AtualizaBenef.Beneficiario(CurrentSystem, CLng(ServiceVar("HANDLEBENEFICIARIO")), qEnderecoBeneficiario.FieldByName("ENDERECORESIDENCIAL").AsInteger, mensagem)

    Set AtualizaBenef = Nothing
	Set qEnderecoBeneficiario = Nothing
End Sub
