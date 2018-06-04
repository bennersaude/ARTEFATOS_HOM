'HASH: 0D36E62442F07937B8BF786648877D30

Public Sub Main
  Dim componente As CSBusinessComponent
  Set componente = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.SamBeneficiarioBLL, Benner.Saude.Beneficiarios.Business")
  componente.AddParameter(pdtString, CStr(SessionVar("CONTRATOS")))
  componente.AddParameter(pdtInteger, CInt(SessionVar("QTDDIAS")))
  componente.Execute("EnvioDocPendentesBeneficiario")
  Set componente = Nothing
End Sub
