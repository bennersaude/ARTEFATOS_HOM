'HASH: A375E893AD13735D93D6F2228D70988D
'#Uses "*addXMLAtt"


Public Sub Main
  	Dim psDigitado As String
	Dim handle As Long
	Dim piRecebedor As Long
	Dim pdDataBase As Date
	Dim psTabelaPreco As String
	Dim viVersaoTISS As Long

	psDigitado = CStr( ServiceVar("psDigitado") )
	piRecebedor = ServiceVar("piRecebedor")
	If CStr( ServiceVar("pdDataBase") ) <> "" Then
		pdDataBase = CDate( ServiceVar("pdDataBase") )
	End If
	psTabelaPreco = CStr( ServiceVar("psTabelaPreco") )

	Dim sqlCount As BPesquisa
    Set sqlCount=NewQuery
	Dim sql As BPesquisa
	Set sql=NewQuery
	Dim sp As BStoredProc
	Set sp=NewStoredProc

	Dim qVersao As BPesquisa
	Set qVersao = NewQuery
	qVersao.Clear
	qVersao.Add("SELECT MAX(HANDLE) HANDLE FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'")
	qVersao.Active = True
	viVersaoTISS = qVersao.FieldByName("HANDLE").AsInteger

	On Error GoTo erro

	sp.Name="BSTISS_VALIDAREVENTO"
	sp.AddParam("P_DIGITADO",ptInput, ftString, 100)
	sp.AddParam("P_TABELAPRECO",ptInput, ftString,30)
	sp.AddParam("P_RECEBEDOR", ptInput,ftInteger)
	sp.AddParam("P_DATABASE", ptInput, ftDateTime)
	sp.AddParam("P_VERSAOTISS", ptInput, ftInteger)
	sp.AddParam("P_VALIDARVIGENCIA",ptInput, ftString,1)
	sp.AddParam("P_GRAURETORNO",ptOutput, ftInteger)
	sp.AddParam("P_HANDLE",ptOutput, ftInteger)

	sp.ParamByName("P_DIGITADO").AsString = psDigitado
	sp.ParamByName("P_TABELAPRECO").AsString= psTabelaPreco
	sp.ParamByName("P_RECEBEDOR").AsInteger= piRecebedor
	sp.ParamByName("P_DATABASE").AsDateTime= pdDataBase
	sp.ParamByName("P_VERSAOTISS").AsInteger = viVersaoTISS
	sp.ParamByName("P_VALIDARVIGENCIA").AsString = "N"
	sp.ExecProc
	handle = sp.ParamByName("P_HANDLE").AsInteger

	Dim xml As String
	Dim vEhEstrutura As Boolean
	Dim vDigitadoSemPonto As String

	If (handle>0) Then
	    sql.Clear
		sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE HANDLE=:HANDLE")

		If psTabelaPreco <> "" Then
			sql.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ") ")
		End If

		sql.ParamByName("HANDLE").AsInteger = handle
		sql.Active=True
	Else
	    vEhEstrutura = False
	    vDigitadoSemPonto = Replace(psDigitado, ".", "")
	    If IsNumeric(vDigitadoSemPonto) Then
	      vEhEstrutura = True
	    End If
	    If vEhEstrutura Then

			If psTabelaPreco <> "" Then

		    	sqlCount.Clear
		    	sqlCount.Add("SELECT COUNT(1) QTD FROM SAM_TGE WHERE ESTRUTURA LIKE '"+psDigitado+"%' AND ULTIMONIVEL='S' ")
		    	sqlCount.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ")")

				sql.Clear
		    	sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE ESTRUTURA LIKE '"+psDigitado+"%' AND ULTIMONIVEL='S' ")
		    	sql.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ")")
		    Else

		    	sqlCount.Clear
		    	sqlCount.Add("SELECT COUNT(1) QTD FROM SAM_TGE WHERE ESTRUTURA LIKE '"+psDigitado+"%' AND ULTIMONIVEL='S' ")

		    	sql.Clear
		    	sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE ESTRUTURA LIKE '"+psDigitado+"%' AND ULTIMONIVEL='S' ")
		    End If

            sqlCount.Active = True
            If sqlCount.FieldByName("QTD").AsInteger > 500 Then
              sql.Add("  AND HANDLE = -1 ")
            End If
		    sql.Active=True
	    Else

			If psTabelaPreco <> "" Then

                sqlCount.Clear
		    	sqlCount.Add("SELECT COUNT(1) QTD FROM SAM_TGE WHERE UPPER(Z_DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")
		    	sqlCount.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ")")

		    	sql.Clear
		    	sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE UPPER(Z_DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")
		    	sql.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ")")
		    Else
				sqlCount.Clear
		        sqlCount.Add("SELECT COUNT(1) QTD FROM SAM_TGE WHERE UPPER(Z_DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")

		    	sql.Clear
		    	sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE UPPER(Z_DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")
		    End If

            sqlCount.Active = True
            If sqlCount.FieldByName("QTD").AsInteger > 1000 Then
              sql.Add("  AND HANDLE = -1 ")
            End If

		    sql.Active=True

		    If sql.FieldByName("ESTRUTURA").AsString = "" Then


		      sqlCount.Clear
		      sql.Clear

 			  If psTabelaPreco <> "" Then
		    	sqlCount.Add("SELECT COUNT(1) QTD FROM SAM_TGE WHERE UPPER(DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")
		    	sqlCount.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ")")

		    	sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE UPPER(DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")
		    	sql.Add(" AND HANDLE IN (SELECT A.EVENTO FROM SAM_TGE_TABELATISS A JOIN TIS_TABELAPRECO B ON (A.TABELATISS = B.HANDLE) WHERE B.CODIGO = '" +psTabelaPreco + "' AND B.VERSAOTISS = " + CStr(viVersaoTISS) + ")")
		      Else
		    	sqlCount.Add("SELECT COUNT(1) QTD FROM SAM_TGE WHERE UPPER(DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")

		    	sql.Add("SELECT ESTRUTURA, DESCRICAO, HANDLE FROM SAM_TGE WHERE UPPER(DESCRICAO) LIKE '"+ UCase(psDigitado) +"%' AND ULTIMONIVEL='S' ")
		      End If

              sqlCount.Active = True
              If sqlCount.FieldByName("QTD").AsInteger > 1000 Then
                sql.Add("  AND HANDLE = -1 ")
              End If

		      sql.Active=True

		    End If
		End If
	End If

	xml="<registros>"
	While Not sql.EOF
		xml=xml + "<registro>"
		xml=xml + addXMLAtt( "handle", "handle", sql, "")
		xml=xml + addXMLAtt( "descricao", "descricao", sql, "caption='Descrição' width='300'")
		xml=xml + addXMLAtt( "codigo", "estrutura", sql, "caption='Código' width='120'")
		xml=xml + addXMLAtt( "descricaoDetalhada", sql.FieldByName("ESTRUTURA").AsString+" - "+sql.FieldByName("DESCRICAO").AsString, Nothing, "")
		xml=xml + "</registro>"
		sql.Next
	Wend
	xml=xml + "</registros>"

	ServiceVar("psResult") = CStr(xml)
	GoTo fim

	erro:

		ServiceVar("psResult") = CStr(Err.Description)

	fim:

	Set sql=Nothing
	Set sqlCount=Nothing
	Set qVersao = Nothing
	Set sp=Nothing
End Sub
