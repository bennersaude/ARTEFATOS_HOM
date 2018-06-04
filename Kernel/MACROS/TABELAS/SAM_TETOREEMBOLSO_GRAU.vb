'HASH: 0281F8D594A69E021EF0A3C3184207AC
'#Uses "*bsShowMessage"

Option Explicit

Dim vHandleRegistro As Long

Public Sub TABLE_AfterScroll()
	vHandleRegistro = CurrentQuery.FieldByName("GRAU").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.TetoReembolso.SamTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

		componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_GRAU")
		componente.AddParameter(pdtString, "GRAU")
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("GRAU").AsInteger)
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		'=======================Verifica se existe em outro teto===============================

		Dim msg As String

	    msg = componente.Execute("VerificaExisteEmOutroTeto")

	    If Len(msg) > 0 Then
			bsShowMessage(msg,"E")
	    	CanContinue = False
	    End If

		'=======================Verifica se existe no mesmo teto===============================


	    If vHandleRegistro <> CurrentQuery.FieldByName("GRAU").AsInteger Then

			componente.ClearParameters
			componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_GRAU")
			componente.AddParameter(pdtString, "GRAU")
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("GRAU").AsInteger)
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		    If componente.Execute("VerificaDuplicidadeMesmoTeto") Then
				bsShowMessage("Grau já cadastrado para este TETO","E")
		    	CanContinue = False
		    End If

		End If

		'======================================================================================

	    Set componente = Nothing

End Sub
