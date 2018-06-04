'HASH: 30E1425B940035DCD088DD6540D46FC6

Public Sub Main
	Dim SQL, UPD As Object
	Set SQL  = NewQuery
	Set UPD  = NewQuery

	SQL.Add(" SELECT B.HANDLE, P.CODIGODEREPASSE FROM SAM_BENEFICIARIO B                                                                   ")
	SQL.Add("   JOIN SAM_BENEFICIARIO_REPASSE P ON P.BENEFICIARIO = B.HANDLE                                                               ")
	SQL.Add("  WHERE P.HANDLE = (SELECT MAX(R.HANDLE) FROM SAM_BENEFICIARIO_REPASSE R                                                      ")

	If (InStr(SQLServer, "MSSQL") > 0) Then
		SQL.Add("                      WHERE (R.DATAFINAL IS NULL OR (CONVERT(DATE, SYSDATETIME()) BETWEEN R.DATAINICIAL AND R.DATAFINAL)) ")
	Else
		SQL.Add("                      WHERE (R.DATAFINAL IS NULL OR (TRUNC(SYSDATE) BETWEEN R.DATAINICIAL AND R.DATAFINAL))               ")
	End If

	SQL.Add("                        AND R.BENEFICIARIO = B.HANDLE)                                                                        ")
	SQL.Add("    AND P.CODIGODEREPASSE <> B.CODIGODEREPASSE                                                                                ")
	SQL.Active = True

	While Not SQL.EOF
		UPD.Clear
		UPD.Add(" UPDATE SAM_BENEFICIARIO SET CODIGODEREPASSE = :CODIGODEREPASSE WHERE HANDLE = :BENEFICIARIO                              ")
		UPD.ParamByName("CODIGODEREPASSE").AsString = SQL.FieldByName("CODIGODEREPASSE").AsString
		UPD.ParamByName("BENEFICIARIO").AsInteger = SQL.FieldByName("HANDLE").AsInteger
		UPD.ExecSQL
		SQL.Next
	Wend

	Set SQL = Nothing
	Set UPD = Nothing
End Sub
