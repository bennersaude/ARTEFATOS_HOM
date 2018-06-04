'HASH: 5C772E8BAE49DD011A0EDD2AC98B22B7
Sub POSJobsDiario
  'Claudemir de Souza - 26/05/2003
  '----------------------------------------
  Dim Interface As Object

  Set Interface = CreateBennerObject("BSPOS001.Rotinas")
  Interface.ExecSP(CurrentSystem)
  Set Interface = Nothing

End Sub
