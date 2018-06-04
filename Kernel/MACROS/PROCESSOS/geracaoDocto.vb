'HASH: 3533B0866CD26C8D7113D6F9830B60D5
Sub Main
  Dim piContaFinanceira As Long
  Dim psXml As String
  Dim psMensagem As String
  Dim piResult As Long
  
  piContaFinanceira = CLng(ServiceVar("piContaFinanceira"))
  psXml = CStr(ServiceVar("psXml"))
  psMensagem = CStr(ServiceVar("psMensagem"))
  
  Dim SAMCONTAFINANCEIRA As Object
  Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")
  
  piResult = SAMCONTAFINANCEIRA.GeracaoDocto(CurrentSystem, _
                                                                    piContaFinanceira, _
                                                                    psXml, _
                                                                    psMensagem)
  
  ServiceVar("psMensagem") = psMensagem
  ServiceResult = piResult
End Sub
