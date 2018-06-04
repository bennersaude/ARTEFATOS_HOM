'HASH: 2B8BCC326CF11F9CFCC0EDC993CBE8BD

Public Sub Main
	Dim ModeloRotArq As CSBusinessComponent
	Set ModeloRotArq = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.AgendamentoRotinaArquivo.AgendamentoRotinaArquivo, Benner.Saude.Financeiro.AgendamentoRotinaArquivo")

	If SessionVar("MODELOROTINAARQ") <> "" Then
		ModeloRotArq.AddParameter(pdtInteger, CInt(SessionVar("MODELOROTINAARQ")))
	Else
		ModeloRotArq.AddParameter(pdtInteger, 0)
	End If
  	ModeloRotArq.Execute("GerarProcessarRotinaArquivo")

  	Set ModeloRotArq = Nothing

End Sub
