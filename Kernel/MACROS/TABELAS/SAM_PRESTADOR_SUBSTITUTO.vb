'HASH: 7F06D3AA27139EBFC344322F598CFB58
'#Uses "*bsShowMessage"
Option Explicit

Public Sub DATAINICIOPUBLICACAOPORTAL_OnChange()
	If Not CurrentQuery.FieldByName("DATAINICIOPUBLICACAOPORTAL").IsNull Then
		CurrentQuery.FieldByName("DATAFIMPUBLICACAOPORTAL").AsDateTime = DateAdd("d", 180, CurrentQuery.FieldByName("DATAINICIOPUBLICACAOPORTAL").AsDateTime)
	End If

End Sub

Public Sub PRESTADORSUBSTITUTO_OnChange()
	Dim validou As Boolean
	Dim funcaoEntidade As CSEntityCall

	Set funcaoEntidade = CurrentEntity.CreateCall("ValidarMunicipiosAtendimentoIguais")
	funcaoEntidade.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
	funcaoEntidade.AddParameter(pdtAutomatic, CurrentQuery.FieldByName("PRESTADORSUBSTITUTO").AsInteger)
	validou = CBool(funcaoEntidade.Execute)

	If Not validou Then
		If (bsShowMessage("Prestador Substituto possui local de atendimento em Município diferente do Prestador que está sendo substituído. Deseja confirmar mesmo assim?", "Q") = vbNo) Then
			CurrentQuery.FieldByName("PRESTADORSUBSTITUTO").Clear
		End If
	End If

	Set funcaoEntidade = Nothing
End Sub

