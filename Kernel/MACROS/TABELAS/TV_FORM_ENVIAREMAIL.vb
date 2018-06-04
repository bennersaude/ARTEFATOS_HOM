'HASH: 9087EC49DD70226AFE187D601168A165
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_AfterInsert()
	Dim componente As CSBusinessComponent
	Dim retorno As String

	CurrentQuery.FieldByName("TIPONOTIFICACAO").AsString = SessionVar("TIPONOTIFICACAO")
	CurrentQuery.FieldByName("REMETENTE").AsString = SessionVar("REMETENTE")
	CurrentQuery.FieldByName("ASSUNTO").AsString = "Solicitação de Documentação Pendente"

	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.TvFormEnviarEmailBLL, Benner.Saude.Prestadores.Business")

	componente.AddParameter(pdtInteger, 4)
	componente.AddParameter(pdtInteger, CLng(SessionVar("HANDLEPROCESSO")))

	CurrentQuery.FieldByName("TEXTOEMAIL").AsString = componente.Execute("BuscarTextoEmail")

	Set componente = Nothing

End Sub

Public Sub TABLE_AfterPost()
	Dim componente As CSBusinessComponent

	Dim vHandleProcesso As Long
	Dim vRemetente As String
	Dim vDestinatario As String
	Dim vAssunto As String
	Dim vTextoEmail As String

	On Error GoTo Exception

	vHandleProcesso = CLng(SessionVar("HANDLEPROCESSO"))
	vRemetente = CurrentQuery.FieldByName("REMETENTE").AsString
	vDestinatario = CurrentQuery.FieldByName("DESTINATARIO").AsString
	vAssunto =  CurrentQuery.FieldByName("ASSUNTO").AsString
	vTextoEmail = CurrentQuery.FieldByName("TEXTOEMAIL").AsString



	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.TvFormEnviarEmailBLL, Benner.Saude.Prestadores.Business")

	componente.AddParameter(pdtInteger, vHandleProcesso)
	componente.AddParameter(pdtString, vRemetente)
	componente.AddParameter(pdtString, vDestinatario)
	componente.AddParameter(pdtString, vAssunto)
	componente.AddParameter(pdtString, vTextoEmail)

	componente.Execute("EnviarEmail")

	componente.ClearParameters

	componente.AddParameter(pdtInteger, vHandleProcesso)
	componente.AddParameter(pdtString, vRemetente)
	componente.AddParameter(pdtString, vDestinatario)
	componente.AddParameter(pdtString, vAssunto)
	componente.AddParameter(pdtString, vTextoEmail)

	componente.Execute("SalvarLogProcessoCredenciamento")

	bsShowMessage("E-mail enviado com sucesso!", "I")


	Set componente = Nothing

	Exception:
    	Set componente = Nothing
    	bsShowMessage(Err.Description, "I")
    	Exit Sub
End Sub

