'HASH: E01A3EF53399298E92DFA76B7962FD40

Public Sub Main
	Dim dllGerador As Object
	Dim caminhoWebApp As String
	Set dllGerador = CreateBennerObject("Benner.Saude.WebGenerator.Gerador")

	StartChronometer

	caminhoWebApp = "\\mga-apl039\BENNER\wes\QUALIDADEAGCORRENTE\QualidadeAgWeb"

	dllGerador.RegerarPaginas(caminhoWebApp)

	Set dllGerador = Nothing

	StopChronometer

	MsgBox("Finalizado. Tempo decorrido: " + FormatDateTime2("hh:nn:ss", Chronometer))
End Sub
