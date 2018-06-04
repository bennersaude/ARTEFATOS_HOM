'HASH: 540AEFF913DE6813EF33CB0C47DCCE36
 
'SMS 39578 - Anderson Lonardoni - 20/04/2005
Public Sub BOTAOGERARPFPRESTADO_OnClick()
  Dim Interface As Object
  Set Interface =CreateBennerObject("BSBen013.Rotinas")
  Interface.GeraEventosPFPrestador(CurrentSystem,RecordHandleOfTable("SAM_CONTRATO"))
  Set Interface =Nothing
End Sub
