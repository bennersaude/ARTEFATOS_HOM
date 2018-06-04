'HASH: 51C480E420F57B22AA895A183DBA20E7

'#uses "*CriaTabelaTemporariaSqlServer"

Option Explicit

'--------------------------------------------------------------------------------------------------------------------------
'  SOMENTE USAR A PARTIR DA SAM_AUTORIZ  ----------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------


Public Sub cancelar


	CriaTabelaTemporariaSqlServer
	Dim mensagem As String
	Dim resultado As Integer

	Dim vHandleAutorizacao As Long
	Dim vHandleProtocoloTransacao As Long

	vHandleAutorizacao = RetornaNumeroAutorizacao
	vHandleProtocoloTransacao = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")

	Dim dll As Object
	Set dll=CreateBennerObject("ca043.autorizacao")

	resultado = dll.cancelarAutorizacao( _
		CurrentSystem, _
		vHandleAutorizacao, _
		vHandleProtocoloTransacao, _
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

Public Function RetornaNumeroAutorizacao As Long

  RetornaNumeroAutorizacao = RecordHandleOfTable("SAM_AUTORIZ")

  If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
    Dim qBuscaHandleAutorizacao As Object
    Set qBuscaHandleAutorizacao  = NewQuery

    qBuscaHandleAutorizacao.Clear
    qBuscaHandleAutorizacao.Add("SELECT AUTORIZACAO FROM SAM_PROTOCOLOTRANSACAOAUTORIZ WHERE HANDLE = :HANDLE")
    qBuscaHandleAutorizacao.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
    qBuscaHandleAutorizacao.Active = True
    RetornaNumeroAutorizacao = qBuscaHandleAutorizacao.FieldByName("AUTORIZACAO").AsInteger

    Set qBuscaHandleAutorizacao  = Nothing
  End If

End Function
