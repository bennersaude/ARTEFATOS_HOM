'HASH: 55E8837FB7337775BF9B815165320B6A
'#uses "*CriaTabelaTemporariaSqlServer"
Option Explicit

Public Sub TABLE_AfterPost()
Dim BSAte009 As Object
Dim SQL      As Object
Dim SQLIns   As Object
Dim vsMsg    As String

	If Not VisibleMode Then
		If WebVisionCode = "NOVOCANC" Then
			Set SQL = NewQuery
			SQL.Clear
			SQL.Add("Select B.EVENTO,                                                           ")
			SQL.Add("       B.QTDSOLICITADA,                                                    ")
			SQL.Add("       C.SENHA,                                                            ")
			SQL.Add("       C.SITUACAO                                                          ")
			SQL.Add("  FROM SAM_AUTORIZ                    A                                    ")
			SQL.Add("  Left Join SAM_AUTORIZ_EVENTOSOLICIT B On (B.AUTORIZACAO = A.HANDLE)      ")
			SQL.Add("  Left Join SAM_AUTORIZ_EVENTOGERADO  C On (C.EVENTOSOLICITADO = A.HANDLE) ")
			SQL.Add(" WHERE A.HANDLE = :AUTORIZ                                                 ")
			SQL.ParamByName("AUTORIZ").Value = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
			SQL.Active = True
			SQL.First
			Set SQLIns = NewQuery
			SQLIns.Clear
			SQLIns.Add("INSERT INTO AUT_AUTORIZEXTERNA_EVENTOS(HANDLE,          ")
			SQLIns.Add("                                       AUTORIZEXTERNA,  ")
			SQLIns.Add("                                       EVENTO,          ")
			SQLIns.Add("                                       QUANTIDADE,      ")
			SQLIns.Add("                                       SENHA,           ")
			SQLIns.Add("                                       SITUACAO)        ")
			SQLIns.Add("                                VALUES(:HANDLE,         ")
			SQLIns.Add("                                       :AUTORIZEXTERNA, ")
			SQLIns.Add("                                       :EVENTO,         ")
			SQLIns.Add("                                       :QUANTIDADE,     ")
			SQLIns.Add("                                       :SENHA,          ")
			SQLIns.Add("                                       :SITUACAO)       ")
			While Not SQL.EOF
				SQLIns.ParamByName("HANDLE").Value         = NewHandle("AUT_AUTORIZEXTERNA_EVENTOS")
				SQLIns.ParamByName("AUTORIZEXTERNA").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
				SQLIns.ParamByName("EVENTO").Value         = SQL.FieldByName("EVENTO").AsInteger
				SQLIns.ParamByName("QUANTIDADE").Value     = SQL.FieldByName("QTDSOLICITADA").AsInteger
				SQLIns.ParamByName("SENHA").Value          = SQL.FieldByName("SENHA").AsString
				SQLIns.ParamByName("SITUACAO").Value       = SQL.FieldByName("SITUACAO").AsString
				SQLIns.ExecSQL
				SQL.Next
			Wend
			Set SQLIns = Nothing
			Set SQL = Nothing
		ElseIf WebVisionCode = "EXEC" Then
			Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
			vsMsg = BSAte009.NovoExec(CurrentSystem,CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger,CurrentQuery.FieldByName("HANDLE").AsInteger)
			If vsMsg <> "" Then
				InfoDescription = vsMsg
				'CanContinue = False
			End If
			Set BSAte009 = Nothing
		End If
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If Not VisibleMode Then
		If WebVisionCode = "TIPOTRANS" Then
			CancelDescription = "Para inserir uma nova transação utilize o link do menu."
			CanContinue = False
		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
Dim SQL      As Object
Dim SQL1     As Object
Dim BSAte009 As Object
Dim UsarDLL  As Boolean
Dim vsMsg    As String

	If (Not VisibleMode) Then
		If WebVisionCode = "NOVOCANC" Then
			If CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger > 0 Then
				Set SQL = NewQuery
				SQL.Clear
				SQL.Add("Select A.BENEFICIARIO,                                                ")
				SQL.Add("       B.EXECUTOR,                                                    ")
				SQL.Add("       B.SOLICITANTE,                                                 ")
				SQL.Add("       B.RECEBEDOR,                                                   ")
				SQL.Add("       B.LOCALEXECUCAO,                                               ")
				SQL.Add("       A.CID,                                                         ")
				SQL.Add("       C.GUIA                                                         ")
				SQL.Add("  FROM SAM_AUTORIZ                    A                               ")
				SQL.Add("  Left Join SAM_AUTORIZ_EVENTOSOLICIT B On (B.AUTORIZACAO = A.HANDLE) ")
				SQL.Add("  Left join SAM_GUIA                  c On (c.AUTORIZACAO = A.HANDLE) ")
				SQL.Add(" WHERE A.HANDLE = :AUTORIZ                                            ")
				SQL.ParamByName("AUTORIZ").Value = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				SQL.Active = True
				CurrentQuery.FieldByName("EXECUTOR").Value      = SQL.FieldByName("EXECUTOR").AsInteger
				CurrentQuery.FieldByName("SOLICITANTE").Value   = SQL.FieldByName("SOLICITANTE").AsInteger
				CurrentQuery.FieldByName("RECEBEDOR").Value     = SQL.FieldByName("RECEBEDOR").AsInteger
				CurrentQuery.FieldByName("LOCALEXECUCAO").Value = SQL.FieldByName("LOCALEXECUCAO").AsInteger
	 			CurrentQuery.FieldByName("BENEFICIARIO").Value  = SQL.FieldByName("BENEFICIARIO").AsInteger
	 			If SQL.FieldByName("CID").AsInteger > 0 Then
	 				CurrentQuery.FieldByName("CID").Value       = SQL.FieldByName("CID").AsInteger
	 			End If
	 			CurrentQuery.FieldByName("GUIA").Value          = SQL.FieldByName("GUIA").AsInteger
	 			Set SQL = Nothing
			Else
				CancelDescription = "Não foi encontrado um tipo de autorização para cancelamento." + Chr(13) + "Entre em contato com a Central de Atendimento."
				CanContinue = False
			End If
		ElseIf (WebVisionCode = "NOVOTRANS") Or (WebVisionCode = "NOVOELE") Or (WebVisionCode = "EXEC") Then
			Set SQL = NewQuery
			SQL.Clear
			SQL.Add("SELECT B.HANDLE,                                                             ")
			SQL.Add("       A.DESCRICAO,                                                          ")
			SQL.Add("       A.EXIGESOLICITANTE,                                                   ")
			SQL.Add("       A.EXIGEEXECUTOR,                                                      ")
			SQL.Add("       A.EXIGERECEBEDOR,                                                     ")
			SQL.Add("       A.EXIGELOCALEXECUCAO                                                  ")
			SQL.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
			SQL.Add("  JOIN SIS_TIPOAUTORIZACAOEXTERNA B ON (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
			SQL.Add(" WHERE PADRAOAUTORIZADOREXTERNO = 'S'                                        ")
			SQL.Add("   AND A.HANDLE = :HANDLE                                                    ")
			SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
			SQL.Active = True
			If SQL.FieldByName("DESCRICAO").AsString = "" Then
				CancelDescription = "Não foram encontrados tipos de autorizações com o código selecionado."
				CanContinue = False
			ElseIf SQL.FieldByName("HANDLE").AsInteger = 110 Then
				CancelDescription = ""
				Set SQL1 = NewQuery
				SQL1.Clear
				SQL1.Add("SELECT AUT.BENEFICIARIO,                                                    ")
				SQL1.Add("       AES.EXECUTOR,                                                        ")
				SQL1.Add("       AES.SOLICITANTE,                                                     ")
				SQL1.Add("       AES.RECEBEDOR,                                                       ")
				SQL1.Add("       AES.LOCALEXECUCAO                                                    ")
				SQL1.Add("  FROM SAM_AUTORIZ               AUT                                        ")
				SQL1.Add("  JOIN SAM_AUTORIZ_EVENTOSOLICIT AES ON (AES.AUTORIZACAO      = AUT.HANDLE) ")
				SQL1.Add("  JOIN SAM_AUTORIZ_EVENTOGERADO  AEG ON (AEG.EVENTOSOLICITADO = AES.HANDLE) ")
				SQL1.Add(" WHERE AES.AUTORIZACAO = :HANDLE                                            ")
				SQL1.Add("   AND AES.SITUACAO = 'A'                                                   ")
				SQL1.Add("   AND NOT EXISTS (SELECT 1                                                 ")
                SQL1.Add("                     FROM SAM_AUTORIZ_EVENTOGERADO X                        ")
                SQL1.Add("                    WHERE X.EVENTOSOLICITADO = AES.HANDLE                   ")
                SQL1.Add("                      AND X.SITUACAO IN ('N','C')                           ")
                SQL1.Add("                  )                                                         ")
				SQL1.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsInteger
				SQL1.Active = True
				If SQL1.FieldByName("EXECUTOR").AsInteger > 0 Then
					If SQL1.FieldByName("EXECUTOR").AsInteger <> CurrentQuery.FieldByName("EXECUTOR").AsInteger Then
						CancelDescription = "O prestador 'Executor' selecionado é diferente do 'Executor' da autorização original."
					End If
				End If
				If SQL1.FieldByName("LOCALEXECUCAO").AsInteger > 0 Then
					If SQL1.FieldByName("LOCALEXECUCAO").AsInteger <> CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger Then
						CancelDescription = "O prestador 'Local de execução' selecionado é diferente do 'Local de execução' da autorização original."
					End If
				End If
				If SQL1.FieldByName("RECEBEDOR").AsInteger > 0 Then
					If SQL1.FieldByName("RECEBEDOR").AsInteger <> CurrentQuery.FieldByName("RECEBEDOR").AsInteger Then
						CancelDescription = "O prestador 'Recebedor' selecionado é diferente do 'Recebedor' da autorização original."
					End If
				End If
				If SQL1.FieldByName("BENEFICIARIO").AsInteger <> CurrentQuery.FieldByName("BENEFICIARIO").AsInteger Then
					CancelDescription = "O Beneficiário selecionado é diferente do Beneficiário da autorização original."
				ElseIf (SQL.FieldByName("EXIGEEXECUTOR").AsString = "S") And (CurrentQuery.FieldByName("EXECUTOR").AsString = "") Then
					CancelDescription = "O campo Executor deve ser preenchido."
				ElseIf (SQL.FieldByName("EXIGERECEBEDOR").AsString = "S") And (CurrentQuery.FieldByName("RECEBEDOR").AsString = "") Then
					CancelDescription = "O campo Recebedor deve ser preenchido."
				ElseIf (SQL.FieldByName("EXIGELOCALEXECUCAO").AsString = "S") And (CurrentQuery.FieldByName("LOCALEXECUCAO").AsString = "") Then
					CancelDescription = "O campo Local de Execução deve ser preenchido."
				End If
				Set SQL1 = Nothing
				If CancelDescription <> "" Then
					CanContinue = False
				End If
				Set SQL1 = Nothing
			Else
				CancelDescription = ""
				If (SQL.FieldByName("EXIGEEXECUTOR").AsString = "S") And (CurrentQuery.FieldByName("EXECUTOR").AsString = "") Then
					CancelDescription = "O campo Executor deve ser preenchido."
				ElseIf (SQL.FieldByName("EXIGESOLICITANTE").AsString = "S") And (CurrentQuery.FieldByName("SOLICITANTE").AsString = "") Then
					CancelDescription = "O campo Solicitante deve ser preenchido."
				ElseIf (SQL.FieldByName("EXIGERECEBEDOR").AsString = "S") And (CurrentQuery.FieldByName("RECEBEDOR").AsString = "") Then
					CancelDescription = "O campo Recebedor deve ser preenchido."
				ElseIf (SQL.FieldByName("EXIGELOCALEXECUCAO").AsString = "S") And (CurrentQuery.FieldByName("LOCALEXECUCAO").AsString = "") Then
					CancelDescription = "O campo Local de Execução deve ser preenchido."
				ElseIf SQL.FieldByName("HANDLE").AsInteger = 130 Then
					If CurrentQuery.FieldByName("NUMEROAUTORIZACAO").AsString = "" Then
						CancelDescription = "O número da autorização deve ser digitado."
					ElseIf CurrentQuery.FieldByName("BENEFICIARIO").AsString = "" Then
						CancelDescription = "Um beneficiário deve ser escolhido."
					End If
				ElseIf CurrentQuery.FieldByName("BENEFICIARIO").AsString = "" Then
					CancelDescription = "Um beneficiário deve ser escolhido."
				End If
				If CancelDescription <> "" Then
					CanContinue = False
				End If
			End If
			Set SQL = Nothing
		End If
	End If
End Sub

Public Sub TABLE_NewRecord()
Dim vsSQL  As Object
Dim viCont As Integer

	If Not VisibleMode Then
		CurrentQuery.FieldByName("USUARIO").AsInteger = CurrentUser
		If WebVisionCode = "NOVOCANC"Then
			' Fixo o tipo de autorização com o handle do tipo de autorização externa do tipo cancelamento
			Set vsSQL = NewQuery
			vsSQL.Clear
			vsSQL.Add("Select A.HANDLE                                                              ")
			vsSQL.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
			vsSQL.Add("  Join SIS_TIPOAUTORIZACAOEXTERNA B On (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
			vsSQL.Add(" WHERE A.PADRAOAUTORIZADOREXTERNO = 'S'                                      ")
			vsSQL.Add("   And A.TIPOAUTORIZACAOEXTERNA = 130                                        ")
			vsSQL.Active = True
			CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger = vsSQL.FieldByName("HANDLE").AsInteger
			Set vsSQL = Nothing
		ElseIf WebVisionCode = "NOVOELE" Then
			' Fixo o tipo de autorização com o handle do tipo de autorização externa do tipo elegibilidade
			Set vsSQL = NewQuery
			vsSQL.Clear
			vsSQL.Add("Select A.HANDLE                                                              ")
			vsSQL.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
			vsSQL.Add("  Join SIS_TIPOAUTORIZACAOEXTERNA B On (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
			vsSQL.Add(" WHERE A.PADRAOAUTORIZADOREXTERNO = 'S'                                      ")
			vsSQL.Add("   And A.TIPOAUTORIZACAOEXTERNA = 160                                        ")
			vsSQL.Active = True
			CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger = vsSQL.FieldByName("HANDLE").AsInteger
			viCont = 0
			While viCont < 4
				vsSQL.Clear
				vsSQL.Add("Select COUNT(1) QTD                 ")
				vsSQL.Add("  FROM Z_GRUPOUSUARIOS_PRESTADORAEX ")
				vsSQL.Add(" WHERE USUARIO = :USUARIO           ")
				Select Case viCont
					Case 0
						vsSQL.Add("   And RECEBEDOR = 'S'      ")
					Case 1
						vsSQL.Add("   And EXECUTOR = 'S'       ")
					Case 2
						vsSQL.Add("   And SOLICITANTE = 'S'    ")
					Case 3
						vsSQL.Add("   And LOCALEXECUCAO = 'S'  ")
				End Select
				vsSQL.ParamByName("USUARIO").Value = CurrentUser
				vsSQL.Active = True
				If vsSQL.FieldByName("QTD").AsInteger = 1 Then
					vsSQL.Clear
					vsSQL.Add("Select PRESTADOR                    ")
					vsSQL.Add("  FROM Z_GRUPOUSUARIOS_PRESTADORAEX ")
					vsSQL.Add(" WHERE USUARIO = :USUARIO           ")
					Select Case viCont
						Case 0
							vsSQL.Add("   And RECEBEDOR = 'S'      ")
						Case 1
							vsSQL.Add("   And EXECUTOR = 'S'       ")
						Case 2
							vsSQL.Add("   And SOLICITANTE = 'S'    ")
						Case 3
							vsSQL.Add("   And LOCALEXECUCAO = 'S'  ")
					End Select
					vsSQL.ParamByName("USUARIO").Value = CurrentUser
					vsSQL.Active = True
					Select Case viCont
						Case 0
							CurrentQuery.FieldByName("RECEBEDOR").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
						Case 1
							CurrentQuery.FieldByName("EXECUTOR").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
						Case 2
							CurrentQuery.FieldByName("SOLICITANTE").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
						Case 3
							CurrentQuery.FieldByName("LOCALEXECUCAO").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
					End Select
				End If
				viCont = viCont + 1
			Wend
			Set vsSQL = Nothing
		ElseIf WebVisionCode = "NOVOFECHA" Then
			' Fixo o tipo de autorização com o handle do tipo de autorização externa do tipo fechamento
			Set vsSQL = NewQuery
			vsSQL.Clear
			vsSQL.Add("Select A.HANDLE                                                              ")
			vsSQL.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
			vsSQL.Add("  Join SIS_TIPOAUTORIZACAOEXTERNA B On (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
			vsSQL.Add(" WHERE A.PADRAOAUTORIZADOREXTERNO = 'S'                                      ")
			vsSQL.Add("   And A.TIPOAUTORIZACAOEXTERNA = 150                                        ")
			vsSQL.Active = True
			CurrentQuery.FieldByName("TIPOAUTORIZACAO").Value = vsSQL.FieldByName("HANDLE").AsInteger
			'Verifico se só existe um prestador associado ao usuário
			vsSQL.Clear
			vsSQL.Add("Select COUNT(1) QTD                 ")
  			vsSQL.Add("  FROM Z_GRUPOUSUARIOS_PRESTADORAEX ")
 			vsSQL.Add(" WHERE RECEBEDOR = 'S'              ")
			vsSQL.Add("   And USUARIO = :USUARIO           ")
			vsSQL.ParamByName("USUARIO").Value = CurrentUser
			vsSQL.Active = True
			If vsSQL.FieldByName("QTD").AsInteger = 1 Then
				vsSQL.Clear
				vsSQL.Add("Select PRESTADOR                    ")
  				vsSQL.Add("  FROM Z_GRUPOUSUARIOS_PRESTADORAEX ")
	 			vsSQL.Add(" WHERE RECEBEDOR = 'S'              ")
				vsSQL.Add("   And USUARIO = :USUARIO           ")
				vsSQL.ParamByName("USUARIO").Value = CurrentUser
				vsSQL.Active = True
				CurrentQuery.FieldByName("RECEBEDOR").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
			End If
			Set vsSQL = Nothing
		ElseIf (WebVisionCode = "NOVOTRANS") Or (WebVisionCode = "NOVOELE") Or (WebVisionCode = "EXEC") Then
			Set vsSQL = NewQuery
			If WebVisionCode = "NOVOTRANS" Then
				vsSQL.Clear
				vsSQL.Add("Select A.HANDLE                                                              ")
				vsSQL.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
				vsSQL.Add("  Join SIS_TIPOAUTORIZACAOEXTERNA B On (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
				vsSQL.Add(" WHERE A.PADRAOAUTORIZADOREXTERNO = 'S'                                      ")
				vsSQL.Add("   And A.TIPOAUTORIZACAOEXTERNA = 100                                        ")
				vsSQL.Active = True
				CurrentQuery.FieldByName("TIPOAUTORIZACAO").Value = vsSQL.FieldByName("HANDLE").AsInteger
			End If			
			viCont = 0
			While viCont < 4
				vsSQL.Clear
				vsSQL.Add("Select COUNT(1) QTD                 ")
				vsSQL.Add("  FROM Z_GRUPOUSUARIOS_PRESTADORAEX ")
				vsSQL.Add(" WHERE USUARIO = :USUARIO           ")
				Select Case viCont
					Case 0
						vsSQL.Add("   And RECEBEDOR = 'S'      ")
					Case 1
						vsSQL.Add("   And EXECUTOR = 'S'       ")
					Case 2
						vsSQL.Add("   And SOLICITANTE = 'S'    ")
					Case 3
						vsSQL.Add("   And LOCALEXECUCAO = 'S'  ")
				End Select
				vsSQL.ParamByName("USUARIO").Value = CurrentUser
				vsSQL.Active = True
				If vsSQL.FieldByName("QTD").AsInteger = 1 Then
					vsSQL.Clear
					vsSQL.Add("Select PRESTADOR                    ")
					vsSQL.Add("  FROM Z_GRUPOUSUARIOS_PRESTADORAEX ")
					vsSQL.Add(" WHERE USUARIO = :USUARIO           ")
					Select Case viCont
						Case 0
							vsSQL.Add("   And RECEBEDOR = 'S'      ")
						Case 1
							vsSQL.Add("   And EXECUTOR = 'S'       ")
						Case 2
							vsSQL.Add("   And SOLICITANTE = 'S'    ")
						Case 3
							vsSQL.Add("   And LOCALEXECUCAO = 'S'  ")
					End Select
					vsSQL.ParamByName("USUARIO").Value = CurrentUser
					vsSQL.Active = True
					Select Case viCont
						Case 0
							CurrentQuery.FieldByName("RECEBEDOR").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
						Case 1
							CurrentQuery.FieldByName("EXECUTOR").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
						Case 2
							CurrentQuery.FieldByName("SOLICITANTE").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
						Case 3
							CurrentQuery.FieldByName("LOCALEXECUCAO").Value = vsSQL.FieldByName("PRESTADOR").AsInteger
					End Select
				End If
				viCont = viCont + 1
			Wend
			Set vsSQL = Nothing
			If WebVisionCode = "EXEC" Then
				Set vsSQL = NewQuery
				vsSQL.Clear
				vsSQL.Add("Select HANDLE                                                                                          ")
	  			vsSQL.Add("FROM SAM_TIPOAUTORIZ                                                                                   ")
				vsSQL.Add(" WHERE PADRAOAUTORIZADOREXTERNO = 'S'                                                                  ")
				vsSQL.Add("   And TIPOAUTORIZACAOEXTERNA = (Select HANDLE FROM SIS_TIPOAUTORIZACAOEXTERNA WHERE CODIGO = :CODIGO) ")
				vsSQL.ParamByName("CODIGO").Value = 110
				vsSQL.Active = True
				If vsSQL.FieldByName("HANDLE").AsInteger > 0 Then
					CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger = vsSQL.FieldByName("HANDLE").AsInteger
				Else
					CancelDescription = "Não foi encontrado um tipo de autorização cadastrada para o tipo 110"
				End If
				Set vsSQL = Nothing
			End If
		End If
	End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
Dim BSAte009 As Object
Dim vsMsg    As String
Dim vsSQL    As Object
Dim vsSql1   As Object
Dim qqq      As String


	'SMS 70094 - Débora Rebello - 23/10/2006 - inicio

	'******* Criação de tabelas temporárias usadas pelo autorizador externo ********
	'Por poder utilizar várias conexões, é preciso verificar, a cada acesso, se, na conexão sendo usada, essas tabelas
	'estão criadas ou não, sendo necessário criá-las caso elas não existem.
	'Foi colocado transação porque estava ocorrendo problemas de se usar uma conexão ao verificar a existência
	'de uma tabela e outra ao criar as tabelas. Assim, verificava-se que a tabela não existia e, ao criar, dava erro
	'de tabela existente.

	If Not InTransaction Then
		StartTransaction
	End If
    If InStr(SQLServer, "MSSQL")>0 Then
      On Error GoTo CriarTabelas
        Set vsSQL = NewQuery
        vsSQL.Clear
        vsSQL.Add("SELECT count(1) FROM #TMP_MENSAGEM")
        vsSQL.Active = True
        GoTo NaoCriarTabelas
      CriarTabelas:
        CriaTabelaTemporariaSqlServer  'SMS 64232 - utilizando função de macro geral até solução em runner
      NaoCriarTabelas:
    End If
	Set vsSQL = Nothing

	If InTransaction Then
		Commit
	End If
	'SMS 70094 - Débora Rebello - 23/10/2006 - fim


	If (Not VisibleMode) Then
		If WebVisionCode = "TIPOTRANS" Then
			If (CommandID = "PROCSOL") Then
				Set vsSQL = NewQuery
				vsSQL.Clear
				vsSQL.Add("SELECT COUNT(1) QTD                     ")
				vsSQL.Add("  FROM AUT_AUTORIZEXTERNA_EVENTOS       ")
				vsSQL.Add(" WHERE AUTORIZEXTERNA = :AUTORIZEXTERNA ")
				vsSQL.ParamByName("AUTORIZEXTERNA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				vsSQL.Active = True
				Set vsSql1 = NewQuery
				vsSql1.Clear
				vsSql1.Add("SELECT B.HANDLE                                                              ")
				vsSql1.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
				vsSql1.Add("  JOIN SIS_TIPOAUTORIZACAOEXTERNA B ON (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
				vsSql1.Add(" WHERE PADRAOAUTORIZADOREXTERNO = 'S'                                        ")
				vsSql1.Add("   AND A.HANDLE = :HANDLE                                                    ")
				vsSql1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
				vsSql1.Active = True
				If (vsSQL.FieldByName("QTD").AsInteger > 0) Or (vsSql1.FieldByName("HANDLE").AsInteger = 130) Or (vsSql1.FieldByName("HANDLE").AsInteger = 150) Then
					Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
					vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,0)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
					End If
					Set BSAte009 = Nothing
				Else
					CancelDescription = "Insira pelo menos um evento na solicitação."
					CanContinue = False
				End If
				Set vsSql1 = Nothing
				Set vsSQL  = Nothing
			ElseIf (CommandID = "ATUA") Then
				Set vsSql1 = NewQuery
				vsSql1.Clear
				vsSql1.Add("SELECT B.HANDLE                                                              ")
				vsSql1.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
				vsSql1.Add("  JOIN SIS_TIPOAUTORIZACAOEXTERNA B ON (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
				vsSql1.Add(" WHERE PADRAOAUTORIZADOREXTERNO = 'S'                                        ")
				vsSql1.Add("   AND A.HANDLE = :HANDLE                                                    ")
				vsSql1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
				vsSql1.Active = True
				If vsSql1.FieldByName("HANDLE").AsInteger <= 110 Then
					' Chamo a DLL com o tipo 110 para atualizar somente...
					Set BSAte009 = CreateBennerObject("BSATE009.Rotinas")
					vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,110)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
						End If
					Set BSAte009 = Nothing
				ElseIf vsSql1.FieldByName("HANDLE").AsInteger = 120 Then
					Set BSAte009 = CreateBennerObject("BSATE009.Rotinas")
					vsMsg = BSAte009.Atualizar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,120)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
					End If
					Set BSAte009 = Nothing
				End If
				Set vsSql1 = Nothing
			End If
		ElseIf WebVisionCode = "NOVOTRANS" Then
			If CommandID = "VALTRANS" Then
				Set vsSQL = NewQuery
				vsSQL.Clear
				vsSQL.Add("SELECT COUNT(1) QTD                     ")
				vsSQL.Add("  FROM AUT_AUTORIZEXTERNA_EVENTOS       ")
				vsSQL.Add(" WHERE AUTORIZEXTERNA = :AUTORIZEXTERNA ")
				vsSQL.ParamByName("AUTORIZEXTERNA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				vsSQL.Active = True
				Set vsSql1 = NewQuery
				vsSql1.Clear
				vsSql1.Add("SELECT B.HANDLE                                                              ")
				vsSql1.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
				vsSql1.Add("  JOIN SIS_TIPOAUTORIZACAOEXTERNA B ON (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
				vsSql1.Add(" WHERE PADRAOAUTORIZADOREXTERNO = 'S'                                        ")
				vsSql1.Add("   AND A.HANDLE = :HANDLE                                                    ")
				vsSql1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
				vsSql1.Active = True
				If (vsSQL.FieldByName("QTD").AsInteger > 0) Or (vsSql1.FieldByName("HANDLE").AsInteger = 130) Or (vsSql1.FieldByName("HANDLE").AsInteger = 150) Then
					Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
					vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,0)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
					End If
					Set BSAte009 = Nothing
				Else
					CancelDescription = "Insira pelo menos um evento na solicitação."
					CanContinue = False
				End If
				Set vsSql1 = Nothing
				Set vsSQL  = Nothing
			ElseIf CommandID = "NOVOATUA" Then
				Set vsSql1 = NewQuery
				vsSql1.Clear
				vsSql1.Add("SELECT B.HANDLE                                                              ")
				vsSql1.Add("  FROM SAM_TIPOAUTORIZ            A                                          ")
				vsSql1.Add("  JOIN SIS_TIPOAUTORIZACAOEXTERNA B ON (B.HANDLE = A.TIPOAUTORIZACAOEXTERNA) ")
				vsSql1.Add(" WHERE PADRAOAUTORIZADOREXTERNO = 'S'                                        ")
				vsSql1.Add("   AND A.HANDLE = :HANDLE                                                    ")
				vsSql1.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TIPOAUTORIZACAO").AsInteger
				vsSql1.Active = True
				CancelDescription = vsSql1.FieldByName("HANDLE").AsString
				If vsSql1.FieldByName("HANDLE").AsInteger <= 110 Then
					' Chamo a DLL com o tipo 110 para atualizar somente...
					Set BSAte009 = CreateBennerObject("BSATE009.Rotinas")
					vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,110)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
						End If
					Set BSAte009 = Nothing
				ElseIf vsSql1.FieldByName("HANDLE").AsInteger = 120 Then
					Set BSAte009 = CreateBennerObject("BSATE009.Rotinas")
					vsMsg = BSAte009.Atualizar(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,120)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
					End If
					Set BSAte009 = Nothing
				End If
				Set vsSql1 = Nothing
			End If
		ElseIf WebVisionCode = "NOVOCANC" Then
			If CommandID = "VALCANC" Then
				Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
				vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,0)
				If vsMsg <> "" Then
					CancelDescription = vsMsg
					CanContinue = False
				End If
				Set BSAte009 = Nothing
			End If
		ElseIf WebVisionCode = "NOVOFECHA" Then
			If CommandID = "VALFECHA" Then
				Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
				vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,0) 'Fechamento(CurrentSystem,1)
				If vsMsg <> "" Then
					CancelDescription = vsMsg
					CanContinue = False
				End If
				Set BSAte009 = Nothing
			End If
		ElseIf WebVisionCode = "NOVOELE" Then
			If CommandID = "VALELE" Then
				Set vsSQL = NewQuery
				vsSQL.Clear
				vsSQL.Add("SELECT COUNT(1) QTD                     ")
				vsSQL.Add("  FROM AUT_AUTORIZEXTERNA_EVENTOS       ")
				vsSQL.Add(" WHERE AUTORIZEXTERNA = :AUTORIZEXTERNA ")
				vsSQL.ParamByName("AUTORIZEXTERNA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				vsSQL.Active = True
				If (vsSQL.FieldByName("QTD").AsInteger > 0) Then
					Set BSAte009 = CreateBennerObject("BSAte009.Rotinas")
					vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,0)
					If vsMsg <> "" Then
						CancelDescription = vsMsg
						CanContinue = False
					End If
					Set BSAte009 = Nothing
				Else
					CancelDescription = "Insira pelo menos um evento na solicitação."
					CanContinue = False
				End If
				Set vsSQL = Nothing
			End If
		ElseIf WebVisionCode = "EXEC" Then
			If CommandID = "VALEXEC" Then
				Set BSAte009 = CreateBennerObject("BSATE009.Rotinas")
				CancelDescription = "Entrou na DLL..."
				vsMsg = BSAte009.GeraAutorizacao(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,110)
				If vsMsg <> "" Then
					CancelDescription = vsMsg
					CanContinue = False
					End If
				Set BSAte009 = Nothing
			End If
		End If
	End If
End Sub
