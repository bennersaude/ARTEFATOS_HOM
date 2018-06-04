'HASH: 34DB12065D7F72FE59094A63CE612B58
Option Explicit

Public Sub Main
	Dim qtdDias As Integer
	Dim qtdDiasString As String
	Dim qParam As Object
	Set qParam = NewQuery
	qParam.Text = "SELECT QTDDIASMANTERLOGS FROM SAM_PARAMETROSCONECTA"
	qParam.Active = True
	qtdDiasString = qParam.FieldByName("QTDDIASMANTERLOGS").AsString
	qParam.Active = False
	Set qParam = Nothing
    If qtdDiasString <> "" Then
		qtdDias = CInt(qtdDiasString)'2

		If (qtdDias > 1) Then
		  Dim qDel As Object
		  Set qDel = NewQuery
		  If InStr(SQLServer, "MSSQL")>0 Then
		    qDel.Text = "DELETE FROM LOG_ALTERACAOREGISTRO WHERE datediff(DAY, DATAHORA, getdate()) > " + CStr(qtdDias)
		  Else
		    qDel.Text = "DELETE FROM LOG_ALTERACAOREGISTRO WHERE (trunc(sysdate) - trunc(DATAHORA)) > " + CStr(qtdDias)
		  End If
		  qDel.ExecSQL
		  Set qDel = Nothing
		Else
		  Err.Raise(1000, "LimparLogAlteracoes", "O valor do parâmetro QuantidadeDias deve ser maior que 1")
		End If
	Else
	  Err.Raise(1000, "LimparLogAlteracoes", "Falta informar o parâmetro QuantidadeDias")
	End If

End Sub
