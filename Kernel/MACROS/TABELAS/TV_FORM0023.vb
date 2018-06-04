'HASH: 6902D2C6623E1E1527F601A31189D58E
'#uses "*CriaTabelaTemporariaSqlServer"

Option Explicit

'--------------------------------------------------------------------------------------------------------------------------
'  SOMENTE USAR A PARTIR DA SAM_AUTORIZ_EVENTOSOLICIT----------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------


Public Sub cancelar

	CriaTabelaTemporariaSqlServer
	Dim mensagem As String
	Dim resultado As Integer

	Dim dll As Object


	Set dll=CreateBennerObject("ca043.autorizacao")

	resultado = dll.cancelarSolicitado( _
		CurrentSystem, _
		RecordHandleOfTable("SAM_AUTORIZ_EVENTOSOLICIT"), _
		CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsInteger, _
		mensagem)
	
	Set dll=Nothing
	If resultado > 0 Then
		InfoDescription = mensagem
	Else
		InfoDescription = "Cancelamento concluído com sucesso"
	End If
End Sub


Public Sub TABLE_AfterPost()
	cancelar
End Sub




