'HASH: 264AF55A75B86849EB4C5D46BCD3307C
'#Uses "*bsShowMessage"

Option Explicit

Dim vHandleRegistro As Long

Public Sub TABLE_AfterScroll()
	vHandleRegistro = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.TetoReembolso.SamTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

		componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_FINATEND")
		componente.AddParameter(pdtString, "FINALIDADEATENDIMENTO")
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger)
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		'=======================Verifica se existe em outro teto===============================

		Dim msg As String

	    msg = componente.Execute("VerificaExisteEmOutroTeto")

	    If Len(msg) > 0 Then
			bsShowMessage(msg,"E")
	    	CanContinue = False
	    End If

		'=======================Verifica se existe no mesmo teto===============================


	    If vHandleRegistro <> CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger Then

			componente.ClearParameters
			componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_FINATEND")
			componente.AddParameter(pdtString, "FINALIDADEATENDIMENTO")
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger)
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		    If componente.Execute("VerificaDuplicidadeMesmoTeto") Then
				bsShowMessage("Finalidade de atendimento já cadastrada para este TETO","E")
		    	CanContinue = False
		    End If

		End If

		'======================================================================================

	    Set componente = Nothing

End Sub
