'HASH: AB26CEBEDAD732F3EA50F50141E8DD58

Public Sub Main

	Dim ExtracaoPublicoAlvo As Object
	Dim qParticipantesPublicoAlvo As Object

	Set ExtracaoPublicoAlvo = CreateBennerObject("Benner.Saude.Web.Clinica.MalaDireta.maladireta")
	Set qParticipantesPublicoAlvo = NewQuery

	qParticipantesPublicoAlvo.Clear
	qParticipantesPublicoAlvo.Add("SELECT HANDLE 						")
	qParticipantesPublicoAlvo.Add("  FROM CLI_EXTRACAOPUBLICOALVO_MALA  ")
	qParticipantesPublicoAlvo.Add(" WHERE SITUACAO = '1'    			")
	qParticipantesPublicoAlvo.Add("   AND DATAHORAPROCESSAMENTO = NULL  ")
	qParticipantesPublicoAlvo.Add("   AND USUARIOPROCESSAMENTO = NULL  	")
    qParticipantesPublicoAlvo.Active = True

	While Not qParticipantesPublicoAlvo.EOF

		ExtracaoPublicoAlvo.Processar(qParticipantesPublicoAlvo.FieldByName("HANDLE").AsInteger,  CurrentUser)

		qParticipantesPublicoAlvo.Next

	Wend

	Set ExtracaoPublicoAlvo = Nothing
	Set qParticipantesPublicoAlvo = Nothing

End Sub
