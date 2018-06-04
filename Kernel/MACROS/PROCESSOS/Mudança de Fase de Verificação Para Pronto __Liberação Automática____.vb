'HASH: E154740A1E06AB5148060E9160845785

Public Sub Main

	Dim pegs As String
	Dim FasePeg As CSEntityCall

  	Set FasePeg = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.PEG.FasePeg, Benner.Saude.Entidades", "LiberarVerificacaoPegsAutomaticoRetornaStringHandles")

	pegs = FasePeg.Execute()

    If (pegs <> "") Then
    	SessionVar("PEGSMUDARFASEVERIFICACAO") = CStr(pegs)

    	Dim dllSamPeg As Object
    	Set dllSamPeg = CreateBennerObject("SamPeg.PEGLote_Processar")
    	dllSamPeg.Exec(CurrentSystem)
	    Set dll = Nothing

	    SessionVar("PEGSMUDARFASEVERIFICACAO") = ""
	Else
		InfoDescription = "Nenhum Peg na fase de verificação para ser mudado de fase"
	End If

	Set FasePeg = Nothing

End Sub
