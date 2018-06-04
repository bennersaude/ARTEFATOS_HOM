'HASH: 37264A18D75E486652421227FF629E25
Public Sub Main

  Dim vDLLOperation As Object
  Set vDLLOperation = CreateBennerObject("Benner.Saude.WSTiss.TissTransmiteMensagem.Operation.TissTransmiteMensagemOperation")

  ServiceResult = vDLLOperation.OperationPortal(CurrentSystem, CStr(ServiceVar("P_XML")), CLng(ServiceVar("P_PRIORIDADE")), "",0,CStr(ServiceVar("P_ORIGEM")))


  Set vDLLOperation = Nothing

End Sub
