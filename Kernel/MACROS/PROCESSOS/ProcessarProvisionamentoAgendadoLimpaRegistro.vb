'HASH: D23F86F063073195F67B329546ECDB51

Public Sub Main

	Dim limpaRegistro As Object
	Set limpaRegistro = NewQuery

		limpaRegistro.Clear
		limpaRegistro.Add("UPDATE SAM_PEG ")
		limpaRegistro.Add("Set PROCESSOPROVISAO = Null")
		limpaRegistro.Add("WHERE PROCESSOPROVISAO Is Not Null")
		limpaRegistro.ExecSQL


End Sub
