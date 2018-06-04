'HASH: 3F2970A3219BD96D62CDA2C6AA29C4CA
Option Explicit
Public Sub Main

	Dim BCImp As CSBusinessComponent
	Set BCImp = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.ImpBenefLeiauteVariavel.ImpBenefLeiauteVariavel, Benner.Saude.Beneficiarios.ImpBenefLeiauteVariavel")

  	BCImp.Execute("Importacao")

  	Set BCImp = Nothing

End Sub
