'HASH: F69E92A52F5A7CDAFF45D64EE5E42DB8

Public Sub Main

	Dim pegs As String
	Dim FasePeg As CSEntityCall

  	Set FasePeg = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.PEG.FasePeg, Benner.Saude.Entidades", "SelecionaPegsDigitadosParaMudarFaseRetornaStringHandles")

	pegs = FasePeg.Execute()

    If (pegs <> "") Then
    	SessionVar("PEGSMUDARFASEVERIFICACAO") = CStr(pegs)

    	Dim dllSamPeg As Object
    	Set dllSamPeg = CreateBennerObject("SamPeg.PEGLote_Processar")
    	dllSamPeg.Exec(CurrentSystem)
	    Set dll = Nothing

	    SessionVar("PEGSMUDARFASEVERIFICACAO") = ""
	Else
	    InfoDescription = "Nenhum Peg na fase de digitação para ser mudado de fase"
	End If

	Set FasePeg = Nothing


End Sub
