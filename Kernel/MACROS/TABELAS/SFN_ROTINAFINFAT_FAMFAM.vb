'HASH: B5CE08BA4386BC241EF8124895DC8AFC
'Macro: SFN_ROTINAFINFAT_FAMFAM
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
    Obj.CancelarFaturamento(CurrentSystem, RecordHandleOfTable("SFN_ROTINAFINFAT"), "FF", CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
  	If bsShowMessage("Cancelar Faturamento?", "Q") = vbYes Then
	    Dim vsMensagemErro As String
    	Dim viRetorno As Long
    	Dim vcContainer As CSDContainer
		Dim vsContrato As String

    	Set vcContainer = NewContainer

	    vcContainer.AddFields("HANDLE:INTEGER")
    	vcContainer.AddFields("OPCAOCANCELAMENTO:STRING")
    	vcContainer.AddFields("HOPCAOCANCELAMENTO:INTEGER")

	    vcContainer.Insert
    	vcContainer.Field("HANDLE").AsInteger             = RecordHandleOfTable("SFN_ROTINAFINFAT")
    	vcContainer.Field("OPCAOCANCELAMENTO").AsString   = "FF"
    	vcContainer.Field("HOPCAOCANCELAMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

		vsContrato = Solver(CurrentQuery.FieldByName("ROTINAFINFATFAM").AsInteger,"SFN_ROTINAFINFAT_FAM","CONTRATO")

	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
    										"BSBen018", _
    										"RotinaFaturamentoBeneficiarios_Cancelar", _
    										"Cancelando família " + Solver(CurrentQuery.FieldByName("FAMILIA").AsInteger,"SAM_FAMILIA","FAMILIA") + " do Contrato " + Solver(CLng(vsContrato),"SAM_CONTRATO","CONTRATO"), _
    										0, _
    										"", _
    										"", _
    										"", _
    										"", _
    										"C", _
    										False, _
    										vsMensagemErro, _
    										vcContainer)
    	If viRetorno = 0 Then
      		bsShowMessage("Processo enviado para execução no servidor!", "I")
    	Else
      		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    	End If
    End If
  End If
  Set Obj = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
	End Select
End Sub
