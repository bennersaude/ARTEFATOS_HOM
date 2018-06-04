'HASH: 07227CBAE030DDA4930DADEC72614305
'#uses "*Biblioteca"
'#Uses "*CriaTabelaTemporariaSqlServer"

Public Sub Main

	On Error GoTo except

    Dim qSituacao As BPesquisa
	Set qSituacao = NewQuery

    CriaTabelaTemporariaSqlServer

    Dim vContainer As CSDContainer
    Set vContainer = NewContainer

    Dim QueryCalculaPreco As BPesquisa
    Set QueryCalculaPreco = NewQuery

    vContainer.AddFields("CODIGO:STRING;DESCRICAO:STRING;VALOR1:STRING;VALOR2:STRING;VALOR3:STRING")

	ApagaTemporaria

    Dim SP_Chave As Long

	SP_Chave = 0
	NewCounter("AUTORIZADORSP",0,1,SP_Chave)

	Dim SP As BStoredProc
	Set SP = NewStoredProc

	SP.AutoMode = True
	SP.Name = "BSRPTPrecoPrestador_PRE"
	SP.AddParam("p_Prestador",ptInput,ftInteger)
	SP.AddParam("p_EventoI",ptInput,ftInteger)
	SP.AddParam("p_EventoF",ptInput,ftInteger)
	SP.AddParam("p_Filtro",ptInput,ftInteger)
	SP.AddParam("p_Convenio",ptInput,ftInteger)
	SP.AddParam("p_DataBase",ptInput,ftDateTime)
	SP.AddParam("p_Cbos",ptInput,ftInteger)
	SP.AddParam("p_Usuario",ptInput,ftInteger)
	SP.AddParam("p_Chave",ptInput,ftInteger)
	SP.AddParam("p_MascaraTge",ptInput,ftInteger)

	Dim qSQL As BPesquisa
	Set qSQL = NewQuery

    qSQL.Add("  SELECT TGE.HANDLE, TGE.ESTRUTURA, TGE.DESCRICAO, TGE.MASCARATGE                                     ")
    qSQL.Add("	  FROM TIS_TABELAPRECO TTP                                                                          ")
    qSQL.Add("	  JOIN SAM_TGE_TABELATISS TT ON (TT.TABELATISS = TTP.HANDLE)                                        ")
    qSQL.Add("	  JOIN SAM_TGE TGE ON (TGE.HANDLE = TT.EVENTO)                                                      ")
    qSQL.Add("	 WHERE TTP.VERSAOTISS = (SELECT HANDLE FROM TIS_VERSAO WHERE VERSAO = (SELECT MAX(VERSAO) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'))               ")
    qSQL.Add("	   AND TGE.ULTIMONIVEL = 'S'                                                                        ")
    If ServerContainer.Field("MASCARATGE").AsInteger > 0 Then
	  qSQL.Add("   AND TGE.MASCARATGE  = :MASCARATGE                                                                ")
    End If
    qSQL.Add("	   AND TGE.HANDLE NOT IN (SELECT R.EVENTO TEXTO                                                     ")
    qSQL.Add("	                          FROM SAM_PRESTADOR_REGRA R                                                ")
    qSQL.Add("	                         WHERE R.PRESTADOR    = :HANDLE_PRESTADOR                                   ")
    qSQL.Add("	                           AND R.EVENTO       = TGE.HANDLE                                          ")
    qSQL.Add("	                           AND R.REGRAEXCECAO = 'E')                                                ")
    qSQL.Add("	   AND 'A' IN (SELECT 'A' TEXTO                                                                     ")
    qSQL.Add("	                 FROM SAM_PRESTADOR_REGRA R                                                         ")
    qSQL.Add("	                WHERE R.EVENTO       = TGE.HANDLE                                                   ")
    qSQL.Add("	                  AND R.REGRAEXCECAO = 'R'                                                          ")
    qSQL.Add("	                  AND R.PRESTADOR    = :HANDLE_PRESTADOR                                            ")
    qSQL.Add("	               UNION                                                                                ")
    qSQL.Add("	               SELECT 'A' TEXTO                                                                     ")
    qSQL.Add("	                 FROM SAM_PRESTADOR_ESPECIALIDADE ESP                                               ")
    qSQL.Add("	                WHERE ESP.PRESTADOR = :HANDLE_PRESTADOR                                             ")
    qSQL.Add("	                  AND EXISTS (SELECT 'A'                                                            ")
    qSQL.Add("	                                FROM SAM_PRESTADOR_ESPECIALIDADEGRP PRE                             ")
    qSQL.Add("	                               WHERE PRE.PRESTADOR              = ESP.PRESTADOR                     ")
    qSQL.Add("	                                 AND PRE.ESPECIALIDADE          = ESP.ESPECIALIDADE                 ")
    qSQL.Add("	                                 AND PRE.PRESTADORESPECIALIDADE = ESP.HANDLE                        ")
    qSQL.Add("	                                 AND PRE.PERMITERECEBER          = 'S'                              ")
    qSQL.Add("	                                 AND (EXISTS(SELECT 'A' TEXTO                                       ")
    qSQL.Add("	                                               FROM SAM_ESPECIALIDADEGRUPO_EXEC E                   ")
    qSQL.Add("	                                              WHERE E.ESPECIALIDADEGRUPO = PRE.ESPECIALIDADEGRUPO   ")
    qSQL.Add("	                                                AND E.EVENTO             = TGE.HANDLE))             ")
    qSQL.Add("	                                 AND (EXISTS(SELECT 'A' TEXTO                                       ")
    qSQL.Add("	                                               FROM SAM_PRESTADOR_ESPECIALIDADEREG C                ")
    qSQL.Add("	                                              WHERE C.PRESTADORESPECIALIDADEGRP = PRE.HANDLE        ")
    qSQL.Add("	                                                AND C.REGIMEATENDIMENTO         = -1) OR            ")
    qSQL.Add("	                                      NOT EXISTS(SELECT 'A' TEXTO                                   ")
    qSQL.Add("	                                                   FROM SAM_PRESTADOR_ESPECIALIDADEREG D            ")
    qSQL.Add("	                                                  WHERE D.PRESTADORESPECIALIDADEGRP = PRE.HANDLE))) ")
    qSQL.Add("	               UNION                                                                                ")
    qSQL.Add("	               SELECT 'A' TEXTO                                                                     ")
    qSQL.Add("	                 FROM SAM_PRESTADOR_ESPECIALIDADE ESP,                                              ")
    qSQL.Add("	                      SAM_ESPECIALIDADEGRUPO      PRE                                               ")
    qSQL.Add("	                WHERE ESP.PRESTADOR     = :HANDLE_PRESTADOR                                         ")
    qSQL.Add("	                  AND ESP.ESPECIALIDADE = PRE.ESPECIALIDADE                                         ")
    qSQL.Add("	                  AND NOT EXISTS (SELECT 'A'                                                        ")
    qSQL.Add("	                                    FROM SAM_PRESTADOR_ESPECIALIDADEGRP C                           ")
    qSQL.Add("	                                   WHERE C.PRESTADOR = ESP.PRESTADOR                                ")
    qSQL.Add("	                                     AND C.PRESTADORESPECIALIDADE = ESP.HANDLE)                     ")
    qSQL.Add("	                  AND (EXISTS (SELECT 'A' TEXTO                                                     ")
    qSQL.Add("	                                 FROM SAM_ESPECIALIDADEGRUPO_EXEC D                                 ")
    qSQL.Add("	                                WHERE D.ESPECIALIDADEGRUPO = PRE.HANDLE                             ")
    qSQL.Add("	                                  AND D.EVENTO             = TGE.HANDLE)))                          ")
    qSQL.Add("	ORDER BY TGE.ESTRUTURA                                                                              ")
	qSQL.ParamByName("HANDLE_PRESTADOR").AsInteger = ServerContainer.Field("PRESTADOR").AsInteger
    If ServerContainer.Field("MASCARATGE").AsInteger > 0 Then
      qSQL.ParamByName("MASCARATGE").AsInteger = ServerContainer.Field("MASCARATGE").AsInteger
    End If
	qSQL.Active = True

	Dim qEvento As BPesquisa
	Set qEvento = NewQuery

	Dim qVerifica As Object
	Set qVerifica = NewQuery

	While Not qSQL.EOF

		SP.ParamByName("p_Prestador").AsInteger = ServerContainer.Field("PRESTADOR").AsInteger
		SP.ParamByName("p_EventoI").AsInteger   = qSQL.FieldByName("HANDLE").AsInteger
		SP.ParamByName("p_EventoF").AsInteger   = qSQL.FieldByName("HANDLE").AsInteger
		SP.ParamByName("p_Filtro").AsInteger    = 0
		SP.ParamByName("p_Convenio").AsInteger  = ServerContainer.Field("CONVENIO").AsInteger
		SP.ParamByName("p_DataBase").AsDateTime = ServerContainer.Field("COMPETENCIA").AsDateTime
		SP.ParamByName("p_Cbos").DataType = ftInteger
		SP.ParamByName("p_Cbos").Clear
		SP.ParamByName("p_Usuario").AsInteger   = CurrentUser
		SP.ParamByName("p_Chave").AsInteger     = SP_Chave
        SP.ParamByName("p_MascaraTge").AsInteger= qSQL.FieldByName("MASCARATGE").AsInteger
		SP.ExecProc

		qSQL.Next

	Wend

	Set qVerifica = Nothing
	Set qEvento = Nothing

	Dim vData As String
    vData = Format(ServerNow(),"ddMMyyyyhhmmss")

    Dim sFilePath As String
	sFilePath = "pre068_" & vData & ".txt"

	QueryCalculaPreco.UniDirectional = True

    QueryCalculaPreco.Add("SELECT TTP.CODIGO CODIGO_TABELA, TTP.DESCRICAO TABELAPRECO, TMP.EVENTO_HANDLE,                         ")
    QueryCalculaPreco.Add("       TMP.ESTRUTURA,                                                                                  ")
    QueryCalculaPreco.Add("       TGE.DESCRICAO EVENTO_DESCRICAO,                                                                 ")
    QueryCalculaPreco.Add("       TGE.PROVIDENCIANAFALTA AUTORIZACAOPREVIA,                                                       ")
    QueryCalculaPreco.Add("       TMP.REGIME_HANDLE,                                                                              ")
    QueryCalculaPreco.Add("       TMP.REGIME_DESCRICAO,                                                                           ")
    QueryCalculaPreco.Add("       TMP.GRAU_HANDLE,                                                                                ")
    QueryCalculaPreco.Add("       TMP.GRAU_GRAU,                                                                                  ")
    QueryCalculaPreco.Add("       TMP.GRAU_DESCRICAO,                                                                             ")
    QueryCalculaPreco.Add("       TMP.CBOSCODIGO,                                                                                 ")
    QueryCalculaPreco.Add("       TMP.ORIGEMCALCULO,                                                                              ")
    QueryCalculaPreco.Add("       TMP.VALOR,                                                                                      ")
    QueryCalculaPreco.Add("       TG.GRAUPRINCIPAL,                                                                               ")
    QueryCalculaPreco.Add("       G.ORIGEMVALOR,                                                                                  ")
    QueryCalculaPreco.Add("      (SELECT COUNT(T.HANDLE)                                                                          ")
    QueryCalculaPreco.Add("         FROM SAM_TGE_GRAU T                                                                           ")
    QueryCalculaPreco.Add("         JOIN SAM_GRAU G ON (G.HANDLE = T.GRAU AND G.ORIGEMVALOR = 2)                                  ")
    QueryCalculaPreco.Add("        WHERE (T.EVENTO = TMP.EVENTO_HANDLE)) QTDAUX                                                   ")
    QueryCalculaPreco.Add("  FROM TMP_PRECOS TMP                                                                                  ")
    QueryCalculaPreco.Add("  JOIN SAM_TGE TGE ON (TMP.EVENTO_HANDLE = TGE.HANDLE)                                                 ")
    QueryCalculaPreco.Add("  JOIN SAM_TGE_GRAU TG ON (TGE.HANDLE = TG.EVENTO AND TMP.GRAU_HANDLE = TG.GRAU)                       ")
    QueryCalculaPreco.Add("  JOIN SAM_GRAU G ON (G.HANDLE = TG.GRAU)                                                              ")
    QueryCalculaPreco.Add("  JOIN SAM_TGE_TABELATISS TT ON (TT.EVENTO = TGE.HANDLE)                                               ")
    QueryCalculaPreco.Add("  JOIN TIS_TABELAPRECO TTP ON (TT.TABELATISS = TTP.HANDLE AND TTP.VERSAOTISS = (SELECT MAX(HANDLE) FROM TIS_VERSAO WHERE ATIVODESKTOP = 'S'))")
    QueryCalculaPreco.Add(" WHERE TMP.USUARIO = :P_USUARIO                                                                        ")
    QueryCalculaPreco.Add("   AND TMP.CHAVE = :P_CHAVE                                                                            ")

    If ServerContainer.Field("MASCARATGE").AsInteger > 0 Then
        QueryCalculaPreco.Add(" AND EVENTO_HANDLE IN (SELECT HANDLE FROM SAM_TGE WHERE MASCARATGE = :MASCARATGE)                  ")
        QueryCalculaPreco.ParamByName("MASCARATGE").AsInteger = ServerContainer.Field("MASCARATGE").AsInteger
    End If

    QueryCalculaPreco.Add(" ORDER BY TTP.CODIGO, TGE.ESTRUTURA , TMP.REGIME_DESCRICAO, CASE ")
    QueryCalculaPreco.Add("    WHEN TG.GRAUPRINCIPAL = 'S' THEN 0                                                                 ")
    QueryCalculaPreco.Add("    ELSE 1 END, GRAU_GRAU,                                                                             ")
    QueryCalculaPreco.Add("     ORIGEMVALOR,   CBOSCODIGO                                                                         ")

	QueryCalculaPreco.ParamByName("P_USUARIO").AsInteger = CurrentUser
	QueryCalculaPreco.ParamByName("P_CHAVE").AsInteger   = SP_Chave
	QueryCalculaPreco.Active  = True

    Open sFilePath For Output As #1


    Dim vQtdCustoOperacional   As Double
    Dim vQtdHonorariosMedicos  As Double
    Dim vQtdFilme              As Double
    Dim valorFilme             As Double
    Dim valorCo                As Double
    Dim valorHonorarios        As Double
    Dim vPorteAnestesico       As Double
    Dim vPorteSala             As Double

    Dim vCbos              As String
    Dim vEstrutura         As String
    Dim vEventoDescricao   As String
    Dim vRegime            As String
    Dim vTabelaUS          As String
    Dim vTabelaHM          As String
    Dim vTabelaCO          As String
    Dim vTabelaFilme       As String
    Dim valorTotal         As String
    Dim vGrauDescricao     As String
    Dim vAutorizacaoPrevia As String
    Dim vQtdAux            As String


    Print #1, 	"Estrutura" &  ";" & _
				"Evento" &  ";" &  _
                "Regime" &  ";" &  _
                "Grau" & ";" & _
                "Cbos" & ";" & _
                "AutorizacaoPrevia" & ";" & _
                "Tabela Honorários Médicos" & ";" & _
				"Valor Honorários Médicos" & ";" &  _
                "Tabela Filme" & ";" & _
				"Valor Filme" &  ";" &  _
                "Tabela Custo Operacional" & ";" & _
				"Valor Custo Operacional" &  ";" &  _
                "PorteAnestesico" &  ";" & _
                "PorteSala" &  ";" & _
                "Quantidade Auxiliar" &  ";" & _
                "Valor Total"


    While Not QueryCalculaPreco.EOF

		vQtdCustoOperacional  = 0
		vQtdHonorariosMedicos = 0
		vQtdFilme             = 0
	    valorFilme            = 0
	    valorCo               = 0
	    valorHonorarios       = 0
	    vPorteAnestesico      = 0
	    vPorteSala            = 0

	    vCbos              = ""
	    vEstrutura         = ""
	    vEventoDescricao   = ""
	    vRegime            = ""
	    vTabelaUS          = ""
	    vTabelaHM          = ""
	    vTabelaCO          = ""
	    vTabelaFilme       = ""
	    valorTotal         = ""
	    vGrauDescricao     = ""
	    vAutorizacaoPrevia = ""
	    vQtdAux            = ""


        vContainer.DeleteAll

        vEstrutura         = QueryCalculaPreco.FieldByName("ESTRUTURA").AsString
    	vEventoDescricao   = QueryCalculaPreco.FieldByName("EVENTO_DESCRICAO").AsString
        vRegime            = QueryCalculaPreco.FieldByName("REGIME_DESCRICAO").AsString
        vCbos              = QueryCalculaPreco.FieldByName("CBOSCODIGO").AsString
        valorTotal         = Replace(QueryCalculaPreco.FieldByName("VALOR").AsString, ".", ",")
        vGrauDescricao     = QueryCalculaPreco.FieldByName("GRAU_DESCRICAO").AsString
	   	vQtdAux            = QueryCalculaPreco.FieldByName("QTDAUX").AsString

		Select Case QueryCalculaPreco.FieldByName("AUTORIZACAOPREVIA").AsString
			Case "R"
				vAutorizacaoPrevia = "S"
			Case Else
				vAutorizacaoPrevia = "N"
		End Select


        Call Decerializar (QueryCalculaPreco.FieldByName("ORIGEMCALCULO").AsString, vContainer)

        If QueryCalculaPreco.FieldByName("ORIGEMVALOR").AsInteger = 1 Then

	    'TABELA DE HONORARIOS MEDICOS
		    If vContainer.Locate("DESCRICAO","Tabela de US") Then
			    vTabelaUS = vContainer.Field("VALOR1").AsString
		    End If

    		If vContainer.Locate("DESCRICAO","Qtd US de Honorarios") Then
	    		vQtdHonorariosMedicos = CDbl(vContainer.Field("VALOR2").AsString) '/ 10
		    	vTabelaHM = vContainer.Field("VALOR1").AsString
		    End If

    		If vContainer.Locate("DESCRICAO","Valor das US de Honorarios") Then
	    		valorHonorarios = CDbl(vContainer.Field("VALOR2").AsString) * vQtdHonorariosMedicos
		    End If

    	'CUSTO OPERACIONAL
    		If vContainer.Locate("DESCRICAO","Tabela de Custo Operacional") Then
	    		vTabelaCO = vContainer.Field("VALOR1").AsString
		    End If

    		If vContainer.Locate("DESCRICAO","Qtd US de Custos Operacionais") Then
	    		vQtdCustoOperacional = CDbl(vContainer.Field("VALOR2").AsString) '/ 100
		    End If

    		If vContainer.Locate("DESCRICAO","Valor das US de Custos Operacionais") Then
	    		valorCo = CDbl(vContainer.Field("VALOR2").AsString) * vQtdCustoOperacional
		    End If

    	'FILME
    		If vContainer.Locate("DESCRICAO","Fator de Filme") Then
	    		vQtdFilme = CDbl(vContainer.Field("VALOR2").AsString) '/ 100
		    End If

    		If vContainer.Locate("DESCRICAO","Tabela de Filme") Then
				If vContainer.Field("VALOR3").AsString <> Null Then
	    			valorFilme = CDbl(vContainer.Field("VALOR3").AsString) * vQtdFilme
	    		End If

		    	vTabelaFilme = vContainer.Field("VALOR1").AsString & " / " & vContainer.Field("VALOR2").AsString
            End If
	    End If


	    If QueryCalculaPreco.FieldByName("ORIGEMVALOR").AsInteger = 3 Then ' tabela de porte anestesico
		    If vContainer.Locate("CODIGO","007") Then
			    vPorteAnestesico = vContainer.Field("VALOR2").AsInteger
		    End If
	    ElseIf QueryCalculaPreco.FieldByName("ORIGEMVALOR").AsInteger = 5 Then ' tabela de porte de sala
		    If vContainer.Locate("CODIGO","008") Then
		    	vPorteSala = vContainer.Field("VALOR2").AsInteger
		    End If
	    End If


        Print #1, 	vEstrutura &  ";" & _
					vEventoDescricao &  ";" &  _
                    vRegime &  ";" &  _
                    vGrauDescricao & ";" & _
                    vCbos & ";" & _
                    vAutorizacaoPrevia & ";" & _
                    vTabelaHM & ";" & _
					"R$" & CStr(valorHonorarios) &  ";" &  _
                    vTabelaFilme & ";" & _
					"R$" & CStr(valorFilme) &  ";" &  _
                    vTabelaCO & ";" & _
					"R$" & CStr(valorCo) &  ";" &  _
                    CStr(vPorteAnestesico) &  ";" & _
                    CStr(vPorteSala) &  ";" & _
                    vQtdAux &  ";" & _
                    "R$" & CStr(valorTotal)


        QueryCalculaPreco.Next

    Wend

	Close #1


	Dim qRelacao As Object
	Set qRelacao = NewQuery

	Dim handleRelacao As Long
    handleRelacao = NewHandle("SAM_RELACAOPRECPREST_ANEXO")

    qRelacao.Add("INSERT INTO SAM_RELACAOPRECPREST_ANEXO      ")
    qRelacao.Add("       (HANDLE, RELACAOPRECPREST )          ")
    qRelacao.Add("       VALUES (:HANDLE, :RELACAOPRECPREST ) ")

	qRelacao.ParamByName("HANDLE").AsInteger = handleRelacao
	qRelacao.ParamByName("RELACAOPRECPREST").AsInteger = ServerContainer.Field("HANDLE").AsInteger
	qRelacao.ExecSQL

    SetFieldDocument("SAM_RELACAOPRECPREST_ANEXO","RELATORIOANEXO", handleRelacao, sFilePath , True)

	qSituacao.Add("UPDATE SAM_RELACAOPRECPREST SET SITUACAO = '5' WHERE HANDLE = :HANDLE")
   	qSituacao.ParamByName("HANDLE").AsInteger = ServerContainer.Field("HANDLE").AsInteger

	qSituacao.ExecSQL

	Set qSituacao = Nothing
    Set qRelacao          = Nothing
	Set vContainer        = Nothing
	Set SP                = Nothing
	Set qSQL              = Nothing
	Set QueryCalculaPreco = Nothing

	Exit Sub

	except :

		qSituacao.Add("UPDATE SAM_RELACAOPRECPREST SET SITUACAO = '1' WHERE HANDLE = :HANDLE")
   		qSituacao.ParamByName("HANDLE").AsInteger = ServerContainer.Field("HANDLE").AsInteger

		qSituacao.ExecSQL
		Set qSituacao = Nothing

		CancelDescription = "Ocorreu o seguinte erro durante a geração do arquivo: " & Err.Description


End Sub

Public Sub Decerializar (OrigemCalculo As String, ByVal vContainer As CSDContainer )

			Dim vArray() As String

			vArray() = Split(OrigemCalculo,"|",-1)

			Dim S As String

			Dim i As Integer

			i = 0

'====== sempre inicia o container para cada evento======
			vContainer.DeleteAll
'=======================================================

			While i <= UBound(vArray) - 1

				S =  vArray(i)

				vContainer.Insert
				vContainer.Field("CODIGO").AsString =  Format(vArray(i),"000")
				vContainer.Field("DESCRICAO").AsString = vArray(i+1)

				If i+2 <= UBound(vArray) Then
					vContainer.Field("VALOR1").AsString = vArray(i+2)
				End If

				If i+3 <= UBound(vArray) Then
					vContainer.Field("VALOR2").AsString = vArray(i+3)
				End If

				If i+4 <= UBound(vArray) Then
					vContainer.Field("VALOR3").AsString = vArray(i+4)
				End If

				i = i + 5
			Wend

			vContainer.OrderBy("CODIGO")

			vContainer.First

End Sub

Public Sub ApagaTemporaria
	Dim query As Object
	Set query = NewQuery

	query.Add("DELETE FROM TMP_PRECORELATORIO WHERE USUARIO = :USUARIO")
	query.ParamByName("USUARIO").AsInteger = CurrentUser
	query.ExecSQL

	query.Clear
	query.Add("DELETE FROM TMP_PRECOS WHERE USUARIO = :USUARIO")
	query.ParamByName("USUARIO").AsInteger = CurrentUser
	query.ExecSQL

	Set query = Nothing
End Sub
