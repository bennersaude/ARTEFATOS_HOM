'HASH: 7EF4BB113096EF766BD878D3D22DFD2C

Public Sub Main
  Dim componente As CSBusinessComponent
  Set componente = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.SamPegAnexoBll, Benner.Saude.ProcessamentoContas.Business")
  componente.AddParameter(pdtString, CStr(SessionVar("DOCUMENTO")))
  componente.Execute("EnviarNotaFiscalParaCorporativo")
  Set componente = Nothing
End Sub
