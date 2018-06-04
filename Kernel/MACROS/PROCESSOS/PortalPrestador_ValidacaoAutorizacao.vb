'HASH: BF568DC7430CD08068A4004900A6DA11
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub Main

On Error GoTo erro

If (InStr(SQLServer, "MSSQL") > 0) Then
    CriaTabelaTemporariaSqlServer
End If

Procedure:
On Error GoTo Erro
Dim sp As Object
Dim viNumeroAutorizacao As Long

NewCounter("SAM_AUTORIZ", 0, 1, viNumeroAutorizacao)

Set sp = NewStoredProc
sp.Name = "BSAUT_AUTORIZWEB"

sp.AddParam("P_VERSAOTISS",ptInput, ftInteger)
sp.AddParam("P_WEBAUTORIZ",ptInput, ftInteger)
sp.AddParam("P_TIPOOPERACAO",ptInput, ftInteger)
sp.AddParam("P_AUTORIZACAO",ptInput, ftInteger)
sp.AddParam("P_TIPOTISS",ptInput, ftString)
sp.AddParam("P_ORIGEM",ptInput, ftString)
sp.AddParam("P_USUARIO",ptInput, ftInteger)
sp.AddParam("P_NUMEROAUTORIZACAO", ptInput, ftFloat)
sp.AddParam("P_EHREEMBOLSO",ptInput, ftString)
sp.AddParam("P_RETORNO",ptOutput, ftString)

sp.ParamByName("P_TIPOOPERACAO").AsInteger = CLng(ServiceVar("P_TIPOOPERACAO"))
sp.ParamByName("P_AUTORIZACAO").AsInteger = CLng(ServiceVar("P_AUTORIZACAO"))
sp.ParamByName("P_TIPOTISS").AsString = CStr(ServiceVar("P_TIPOTISS"))
sp.ParamByName("P_ORIGEM").AsString = CStr(ServiceVar("P_ORIGEM"))
sp.ParamByName("P_VERSAOTISS").AsInteger = CLng(ServiceVar("P_VERSAOTISS"))
sp.ParamByName("P_USUARIO").AsInteger = CLng(ServiceVar("P_USUARIO"))
sp.ParamByName("P_WEBAUTORIZ").AsInteger = CLng(ServiceVar("P_WEBAUTORIZ"))
sp.ParamByName("P_NUMEROAUTORIZACAO").AsInteger = viNumeroAutorizacao

sp.ExecProc

If (sp.ParamByName("P_RETORNO").AsString <> "")Then
	ServiceResult = sp.ParamByName("P_RETORNO").AsString
Else
	ServiceResult = ""
End If

Exit Sub

Erro:
    InfoDescription = Err.Description
    CancelDescription = Err.Description
    If WebMode Then
      ServiceResult = Err.Description
    End If


End Sub
