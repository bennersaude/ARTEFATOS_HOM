'HASH: A1952B4B91FA4344E3936671D7D4CAD2

Public Sub Main
	Dim business As CSBusinessComponent
	Set business = BusinessComponent.CreateInstance("Benner.Saude.eSocial.Business.ProcessoAgendado, Benner.Saude.eSocial.Business")
	business.Execute("Processar")
End Sub
