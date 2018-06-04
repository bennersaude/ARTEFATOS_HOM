'HASH: D841199EAE1DF461F165710FD5BC679B

Public Sub Main

	Dim qBuscaRotinaExtracao As Object

	Set qBuscaRotinaExtracao = NewQuery
	qBuscaRotinaExtracao.Clear
	qBuscaRotinaExtracao.Add("SELECT HANDLE       						")
	qBuscaRotinaExtracao.Add("  FROM CLI_ROTINA_EXTRACAOPUBLICOALVO     ")
	qBuscaRotinaExtracao.Add(" WHERE SITUACAO = '1'    					")
	qBuscaRotinaExtracao.Add("   AND DATAHORAPROCESSAMENTO = NULL  		")
	qBuscaRotinaExtracao.Add("   AND USUARIOPROCESSAMENTO = NULL  		")
    qBuscaRotinaExtracao.Active = True

    While Not qBuscaRotinaExtracao.EOF

		Set sp = NewStoredProc
		sp.Name = "BS_000AFB8E"

		sp.AddParam("p_handleRotina",ptInput, ftInteger)
		sp.AddParam("p_usuarioProcessamento",ptInput, ftInteger)

		sp.ParamByName("p_handleRotina").AsInteger = qBuscaRotinaExtracao.FieldByName("HANDLE").AsInteger
		sp.ParamByName("p_usuarioProcessamento").AsInteger = CurrentUser

		sp.ExecProc

        qBuscaRotinaExtracao.Next

    Wend

    Set qBuscaRotinaExtracao = Nothing

End Sub
