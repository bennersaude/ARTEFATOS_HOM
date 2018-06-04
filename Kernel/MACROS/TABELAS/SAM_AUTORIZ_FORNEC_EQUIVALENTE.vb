'HASH: B7E3BDC325157F37DC35BA68284E9544
 
'Macro: SAM_AUTORIZ_FORNEC_EQUIVALENTE

Public Sub TABLE_AfterScroll()
	Dim qTGE As BPesquisa
	Set qTGE = NewQuery

	qTGE.Active = False
	qTGE.Clear
	qTGE.Add(" SELECT E.ESTRUTURA, E.NUMEROREGISTRO, E.DESCRICAOABREVIADA, G.DESCRICAO,                     ")
	qTGE.Add("        E.DESCRICAO DESCRICAOTGE, E.UNIDADE, F.DESCRICAO FORNECEDORPESSOA, E.VIAADMINISTRACAO ")
	qTGE.Add("   FROM SAM_TGE E                                                                             ")
	qTGE.Add("   LEFT JOIN SAM_MATMEDGRUPO G ON E.GRUPOFARMACOLOGICO = G.HANDLE                             ")
	qTGE.Add("   LEFT JOIN SAM_MATMEDBRFORNECEDOR F ON E.FORNECEDORPESSOA = F.HANDLE                        ")
	qTGE.Add("  WHERE E.HANDLE = :HANDLE                                                                    ")
	qTGE.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("MEDICAMENTOEQUIVALENTE").AsInteger
	qTGE.Active = True

	If Not qTGE.EOF Then
		ROTFABRICANTE.Text = "Fabricante: " + qTGE.FieldByName("FORNECEDORPESSOA").AsString
		ROTMEDICAMENTO.Text = "Medicamento: " + qTGE.FieldByName("ESTRUTURA").AsString + " - " + qTGE.FieldByName("DESCRICAOTGE").AsString
		ROTPRINCIPIOATIVO.Text = "Princípio Ativo: " + qTGE.FieldByName("DESCRICAOABREVIADA").AsString
		ROTREGISTROANVISA.Text = "Registro Anvisa/GGREM: " + qTGE.FieldByName("NUMEROREGISTRO").AsString
	End If

	Set qTGE = Nothing
End Sub
