'HASH: 9CD7D9AAAE310DD593FBE2AFAE2379DA
'#Uses "*bsShowMessage"
Option Explicit

Public Sub BOTAOGERARTPAUTORIZ_OnClick()
  Dim InclusaoExclusaoTipoAutorizacao As CSBusinessComponent
  Set InclusaoExclusaoTipoAutorizacao = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.Business.SamAlertaBenefBLL, Benner.Saude.Beneficiarios.Business")

  InclusaoExclusaoTipoAutorizacao.AddParameter(pdtInteger, RecordHandleOfTable("SAM_ALERTABENEF"))
  InclusaoExclusaoTipoAutorizacao.Execute("AbrirTelaInclusaoExclusaoTipoAutorizacao")

  Set InclusaoExclusaoTipoAutorizacao = Nothing

End Sub
