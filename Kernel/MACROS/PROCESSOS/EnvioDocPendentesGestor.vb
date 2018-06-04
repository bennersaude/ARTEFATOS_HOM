'HASH: A94EB034CAD1141B7BDD3F395E3AF8AF

Public Sub Main
  Dim componente As CSBusinessComponent
  Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.SamBeneficiarioBLL, Benner.Saude.Beneficiarios.Business")
  componente.AddParameter(pdtString, CStr(SessionVar("CONTRATOS")))
  componente.AddParameter(pdtInteger, CInt(SessionVar("QTDDIAS")))
  componente.Execute("EnvioDocPendentesGestor")
  Set componente = Nothing
End Sub
