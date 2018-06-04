'HASH: 4FB4DCA585514C8D52CB6AE498F9F8AC
'#Uses "*CriaTabelaTemporariaSqlServer"
Public Sub Main
 On Error GoTo erro

 Dim spProc As Object
 Dim vsNomeProc As String
 Dim vsParam As String
 Dim vsCampo As String
 Dim vsTipo As String
 Dim vsValor As String

 'vsParam = "CHAVE|1|1234|USUARIO|2|124|"

 If Not InTransaction Then
    StartTransaction
 End If

 If InStr(SQLServer, "SQL") > 0 Then
	Dim SQLx As Object
	Set SQLx = NewQuery

	On Error GoTo TabelasTemporarias

	SQLx.Clear
	SQLx.Add("SELECT 1 FROM #TMP_LIMITE")
	SQLx.Active = True

	Set SQLx = Nothing

	GoTo Procedure

	TabelasTemporarias:
	CriaTabelaTemporariaSqlServer

	Set SQLx = Nothing
    End If

Procedure:
On Error GoTo Erro


 vsNomeProc = ServiceVar("NOMEPROC")
 vsParam = ServiceVar("PARAMS")

 Set spProc = NewStoredProc
 spProc.Name = vsNomeProc

 While (InStr(vsParam, "|") > 0)
   vsCampo = pegaParam(vsParam)
   spProc.AddParam(vsCampo, ptInput)
   vsParam = retornaString(vsParam)
   vsTipo = pegaParam(vsParam)
   vsParam = retornaString(vsParam)
   vsValor = pegaParam(vsParam)
   Select Case (CInt(vsTipo))
   		Case 1 'Integer
   		  spProc.ParamByName(vsCampo).AsInteger = CInt(vsValor)
		Case 2 'Float, Money, Double
		  spProc.ParamByName(vsCampo).AsFloat = CDbl(vsValor)
		Case 3 'String
		  spProc.ParamByName(vsCampo).AsString = vsValor
		Case 4 'DateTime
		  spProc.ParamByName(vsCampo).AsDateTime = CDate(vsValor)
   End Select
   vsParam = retornaString(vsParam)
 Wend

 spProc.ExecProc

 If(InTransaction)Then
     Commit
     ServiceResult = "ok"
 End If

 Set spProc = Nothing

 Exit Sub

Erro:
    InfoDescription = Err.Description
    CancelDescription = Err.Description
    ServiceResult = Err.Description

    If InTransaction Then
    	Rollback
	End If

End Sub
Function retornaString(Param As String) As String
  retornaString = Mid(Param, InStr(Param, "|")+1, Len(Param))
End Function
Function pegaParam(Param As String) As String
	pegaParam = Mid(Param, 1, InStr(Param, "|")-1)
End Function
