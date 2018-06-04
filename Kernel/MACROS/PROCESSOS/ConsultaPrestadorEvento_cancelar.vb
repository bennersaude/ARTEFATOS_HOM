'HASH: 17911D9254F2162E4E17C862FEA669E5

Public Sub Main

 	' Acessada da interface Web


	Dim SamConsultaDLL As Object
	Set SamConsultaDLL = CreateBennerObject("SAMCONSULTA.Consulta")
	SamConsultaDLL.CancelarPrestadorEvento(CurrentSystem)
	Set SamConsultaDLL = Nothing
	

End Sub
