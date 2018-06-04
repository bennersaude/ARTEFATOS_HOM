﻿'HASH: A6BAEFD374F6E1A76EC104895EB2C2A8
Option Explicit

'#Uses "*ProcuraEvento"
'#Uses "*bsShowMessage"

Dim vHandle As Long
Dim vHandleRegistro As Long


Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraEvento(False, EVENTO.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("EVENTO").Value = vHandle
  End If
End Sub

Public Sub TABLE_AfterScroll()
	vHandleRegistro = CurrentQuery.FieldByName("EVENTO").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
		Dim componente As CSBusinessComponent
		Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.TetoReembolso.SamTetoReembolsoBLL, Benner.Saude.Beneficiarios.Business")

		componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_EVENTO")
		componente.AddParameter(pdtString, "EVENTO")
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("EVENTO").AsInteger)
		componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		'=======================Verifica se existe em outro teto===============================

		Dim msg As String

	    msg = componente.Execute("VerificaExisteEmOutroTeto")

	    If Len(msg) > 0 Then
			bsShowMessage(msg,"E")
	    	CanContinue = False
	    End If

		'=======================Verifica se existe no mesmo teto===============================


	    If vHandleRegistro <> CurrentQuery.FieldByName("EVENTO").AsInteger Then

			componente.ClearParameters
			componente.AddParameter(pdtString, "SAM_TETOREEMBOLSO_EVENTO")
			componente.AddParameter(pdtString, "EVENTO")
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("EVENTO").AsInteger)
			componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("TETOREEMBOLSO").AsInteger)

		    If componente.Execute("VerificaDuplicidadeMesmoTeto") Then
				bsShowMessage("Evento já cadastrado para este TETO!","E")
		    	CanContinue = False
		    End If

		End If

		'======================================================================================

	    Set componente = Nothing

End Sub
