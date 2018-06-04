'HASH: 55BC8037C71858D92FBE056F3F70B87C
'#Uses "*bsShowMessage"

Option Explicit

Dim vHandleRegistro As Long

Public Sub TABLE_AfterScroll()
	vHandleRegistro = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.TetoReembolso.SamTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

		componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_LOCALATEND")
		componente.AddParameter(pdtString, "LOCALATENDIMENTO")
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger)
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		'=======================Verifica se existe em outro teto===============================

		Dim msg As String

	    msg = componente.Execute("VerificaExisteEmOutroTeto")

	    If Len(msg) > 0 Then
			bsShowMessage(msg,"E")
	    	CanContinue = False
	    End If

		'=======================Verifica se existe no mesmo teto===============================


	    If vHandleRegistro <> CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger Then

			componente.ClearParameters
			componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_LOCALATEND")
			componente.AddParameter(pdtString, "LOCALATENDIMENTO")
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger)
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		    If componente.Execute("VerificaDuplicidadeMesmoTeto") Then
				bsShowMessage("Local de atendimetno já cadastrado para este TETO!","E")
		    	CanContinue = False
		    End If

		End If

		'======================================================================================

	    Set componente = Nothing

End Sub
