'HASH: B1C1F897DFE9038C8D57431529A29E78
'#Uses "*bsShowMessage"

Option Explicit

Dim vHandleRegistro As Long

Public Sub TABLE_AfterScroll()
	vHandleRegistro = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.TetoReembolso.SamTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

		componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_REGIMEATEND")
		componente.AddParameter(pdtString, "REGIMEATENDIMENTO")
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger)
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		'=======================Verifica se existe em outro teto===============================

		Dim msg As String

	    msg = componente.Execute("VerificaExisteEmOutroTeto")

	    If Len(msg) > 0 Then
			bsShowMessage(msg,"E")
	    	CanContinue = False
	    End If

		'=======================Verifica se existe no mesmo teto===============================


	    If vHandleRegistro <> CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger Then

			componente.ClearParameters
			componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_REGIMEATEND")
			componente.AddParameter(pdtString, "REGIMEATENDIMENTO")
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger)
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		    If componente.Execute("VerificaDuplicidadeMesmoTeto") Then
				bsShowMessage("Regime de Atendimento já cadastrado para este TETO!","E")
		    	CanContinue = False
		    End If

		End If

		'======================================================================================

	    Set componente = Nothing

End Sub
