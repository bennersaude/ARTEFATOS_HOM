'HASH: 9F52E40687FBC648B2796031D6492165
'Macro da tabela SAM_WEBSERVICE_LOG

Option Explicit

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

    If (UCase(CommandID) = "K_BRC_WEBSERVICE_GRAVALOG") Then

 		On Error GoTo Finalizacao

	    Dim qInsert As Object
		Set qInsert = NewQuery

	    qInsert.Add("INSERT INTO SAM_WEBSERVICE_LOG (")
	    qInsert.Add("            HANDLE,             ")
	    qInsert.Add("            DATA,               ")
	    qInsert.Add("            SERVICO,            ")
	    qInsert.Add("            METODO,             ")
	    qInsert.Add("            MENSAGEM)           ")
	    qInsert.Add("     VALUES (                   ")
	    qInsert.Add("            :HANDLE,            ")
	    qInsert.Add("            :DATA,              ")
	    qInsert.Add("            :SERVICO,           ")
	    qInsert.Add("            :METODO,            ")
	    qInsert.Add("            :MENSAGEM)          ")

	    qInsert.ParamByName("HANDLE"  ).AsInteger  = NewHandle("SAM_WEBSERVICE_LOG")
	    qInsert.ParamByName("DATA"    ).AsDateTime = ServerNow
	    qInsert.ParamByName("SERVICO" ).AsString   = Mid(CurrentEntity.TransitoryVars("WS_LOG_Servico").AsString, 1, 4)
	    qInsert.ParamByName("METODO"  ).AsString   = Mid(CurrentEntity.TransitoryVars("WS_LOG_Modulo").AsString, 1, 30)
	    qInsert.ParamByName("MENSAGEM").AsString   = Mid(CurrentEntity.TransitoryVars("WS_LOG_Mensagem").AsString, 1, 250)

	    qInsert.ExecSQL

Finalizacao:

        Set qInsert = Nothing

	    If (Err.Description = "") Then
	        CurrentEntity.TransitoryVars("WS_LOG_Retorno").AsString = ""
        Else
            CurrentEntity.TransitoryVars("WS_LOG_Retorno").AsString = " *Erro: " + Err.Description
	    End If

	End If

End Sub
