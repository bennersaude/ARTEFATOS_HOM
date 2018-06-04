'HASH: 47863B75597E86F9FF1F897E1895F100
 

Public Sub BOTAOGERAEVENTOS_OnClick()
  Dim Interface As Object
  Set Interface =CreateBennerObject("SamKrnl.DupEventoRegraExcessao")
  Interface.GeraEventos(CurrentSystem,RecordHandleOfTable("SAM_PRESTADOR"))
End Sub
