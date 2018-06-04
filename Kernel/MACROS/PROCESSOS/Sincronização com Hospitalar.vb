'HASH: 99C6CA5E12CDECAC5E68777ECB38CC71

Public Sub Main
	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Adm.Businnes.SamSincronizacaoHospitalarBLL, Benner.Saude.Adm.Businnes")
	componente.AddParameter(pdtString, CStr(SessionVar("TIPO")))
	componente.AddParameter(pdtString, CStr(SessionVar("REGISTRO")))
	componente.AddParameter(pdtString, CStr(SessionVar("NOTIFICARUSUARIO")))
	componente.Execute("Sincronizar")
	Set componente = Nothing
End Sub
