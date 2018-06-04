'HASH: 4D06965B83A617A3A2BDAC63DE664C56
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim msg As String
	Dim componente As CSBusinessComponent

	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.TetoReembolso.SamTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

	'==============verificando se está alterando o tipo e o teto já está em algum contrato===================

	If (CurrentQuery.FieldByName("TABTIPO").OldValue <> CurrentQuery.FieldByName("TABTIPO").NewValue) Then

		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)

	    msg = componente.Execute("VerificaExisteTetoContrato")

	    If (msg <> "") Then
			bsShowMessage("Alteração do TIPO do TETO não permitida, pois o Teto já está parametrizado no(s) seguinte(s) contrato(s):" & vbNewLine + msg,"E")
			CanContinue = False
			Set componente = Nothing
			Exit Sub
		End If

	End If

	'============== verifica se o tetos são iguais ==========================================================

	componente.ClearParameters
	componente.AddParameter(pdtString, "")
	componente.AddParameter(pdtString, "")
	componente.AddParameter(pdtInteger, 0)
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)

    msg = componente.Execute("VerificaExisteEmOutroTeto")

    If Len(msg) > 0 Then
		bsShowMessage(msg,"E")
    	CanContinue = False
    End If

    Set componente = Nothing

End Sub
