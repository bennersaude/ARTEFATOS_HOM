'HASH: CC8BF6D61357E23A2DFC8BDFA88FC691
'#uses "*CriaTabelaTemporariaSqlServer"
'#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterPost()
	salvar
	If InTransaction Then
	  Commit
	  StartTransaction
	End If
	revalidar
End Sub
 
Public Sub salvar
	Dim sql As BPesquisa
	Set sql=NewQuery
	sql.Add("UPDATE SAM_AUTORIZ_EVENTOSOLICIT SET EHTRATAMENTOSERIADO='S', PERIODOCONTROLSERIE=:P, QTDPORPERIODO=:Q WHERE HANDLE=:H")
	sql.ParamByName("H").AsInteger=RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT")
	sql.ParamByName("P").AsString=CurrentQuery.FieldByName("PERIODO").AsString
	sql.ParamByName("Q").AsInteger=CurrentQuery.FieldByName("QTDPORPERIODO").AsInteger
	sql.ExecSQL
	Set sql=Nothing

End Sub


Public Sub revalidar
	CriaTabelaTemporariaSqlServer

	Dim retorno As Integer
	Dim mensagem As String
	Dim alertas As String

	Dim dll As Object
	Set dll=CreateBennerObject("ca043.autorizacao")
	retorno = dll.revalidarSolicitado(CurrentSystem, RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT"), alertas, mensagem)
	Set dll=Nothing
	If mensagem <> "" Then
		bsShowMessage(mensagem, "I")
	Else
		SessionVar("alertas") = alertas
		bsShowMessage("Revalidação concluída com sucesso", "I")
	End If

End Sub
