'HASH: 9181AC54E2A9EE3C5A570B184A4D77C1

'Macro SAM_CONTRATO_PLANO
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOTRANSFERIR_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("BSBEN017.Plano")
  interface.TransferirPlano(CurrentSystem, CurrentQuery.FieldByName("PLANO").AsInteger, CurrentQuery.FieldByName("CONTRATO").AsInteger)
  Set interface = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.FieldByName("DATAINCLUSAO").AsDateTime < CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("A data inicial da vigência não pode ser menor que a data de adesão do plano", "E")
    CanContinue = False
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID = "BOTAOTRANSFERIR" Then
		BOTAOTRANSFERIR_OnClick
	End If
End Sub
