'HASH: 0470617FA36CFD17CDC30FC2801227B9
Sub Main
  Dim SAMDEVGUIA As Object
  Dim psXml As String
  Dim psMensagem As String
  Dim piResult As Long
  psXml = CStr(ServiceVar("psXml"))
  Set SAMDEVGUIA = CreateBennerObject("SAMDEVOLUCAOGUIA.Rotinas")
  
  piResult = SAMDEVGUIA.DevolverPeg(psXml, psMensagem)
                                    
  ServiceVar("psMensagem") = psMensagem
  ServiceResult = piResult
End Sub
