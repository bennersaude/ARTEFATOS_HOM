'HASH: CC44708623A673ED67F566668A8D30AB
 
Public Sub SendWorkflowMonitorReport(emailDestinatarios As String, emailRemetente As String)
 Dim emailBody As String 
 emailBody = "Relatório de monitoramento do Workflow " + CStr(ServerNow) + Chr(13) + Chr(10) + Chr(13) + Chr(10) 
 
 Dim Q As BPesquisa 
 Set Q = NewQuery 
 
 inicio = ServerNow 
 
 ' - Quantidade de mensagens pendentes para processamento neste instante. (não inclui macros vba) 
 Q.Clear() 
 Q.Add("SELECT COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (SELECT HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_EVENTS$') AND ESTADO = 1") 
 'Q.ParamByName("MODELOINSTANCIA").AsInteger = CurrentQuery.FieldByName("MODELOINSTANCIA").AsInteger 
 Q.Active = True 
 msgPendentes = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Mensagens pendentes..........................: " + CStr(msgPendentes) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 inicio = ServerNow 
 
 ' - Quantidade de mensagens em processamento neste instante. (não inclui macros vba) 
 Q.Clear() 
 Q.Add("Select COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (Select HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_EVENTS$') AND ESTADO = 2") 
 Q.Active = True 
 msgEmProcessamento = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Mensagens em processamento...................: " + CStr(msgEmProcessamento) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' - Quantidade de mensagens processadas na última hora. (não inclui macros vba) 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("Select COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (Select HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_EVENTS$') AND ESTADO = 3 AND DATAHORA > (SYSDATE-1/24)") 
 Else 
 	Q.Add("Select COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (Select HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_EVENTS$') AND ESTADO = 3 AND DATAHORA > DATEADD(hour,-1, GETDATE())") 
 End If 
 Q.Active = True 
 msgProcessadasUltimaHora = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Mensagens processadas na última hora.........: " + CStr(msgProcessadasUltimaHora) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 inicio = ServerNow 
 
 ' - Quantidade de macros VBA pendentes de execução (fila de macros vba) 
 Q.Clear() 
 Q.Add("SELECT COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (SELECT HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_VBAPROCESS$') AND ESTADO = 1") 
 Q.Active = True 
 msgVbaPendentes = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Macros VBA pendentes.........................: " + CStr(msgVbaPendentes) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' - Quantidade de macros VBA em processamento neste instante (fila de macros vba) 
 Q.Clear() 
 Q.Add("SELECT COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (SELECT HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_VBAPROCESS$') AND ESTADO = 1") 
 Q.Active = True 
 msgVbaProcessando = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Macros VBA em processamento..................: " + CStr(msgVbaProcessando) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' - Quantidade de macros VBA processadas na última hora (fila de macros vba) 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("SELECT COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (SELECT HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_VBAPROCESS$') AND ESTADO = 3 AND DATAHORA > (SYSDATE-1/24)") 
 Else 
	Q.Add("SELECT COUNT(1) TOTAL FROM Z_FILAMENSAGENS WHERE FILA = (SELECT HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_VBAPROCESS$') AND ESTADO = 3 AND DATAHORA > DATEADD(hour,-1, GETDATE())") 
 End If 
 Q.Active = True 
 msgVbaProcessadasUltimaHora = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Macros VBA processadas na ultima hora........: " + CStr(msgVbaProcessadasUltimaHora) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' - Quantidade de fluxos iniciados na última hora 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("select count( * ) TOTAL from z_wfmodeloinstancias where inicio >= (SYSDATE-1/24)") 
 Else 
	Q.Add("select count( * ) TOTAL from z_wfmodeloinstancias where inicio >= DATEADD(hour,-1, GETDATE())") 
 End If 
 Q.Active = True 
 qtdeFluxosCriadosUltimaHora = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Fluxos criados na última hora................: " + CStr(qtdeFluxosCriadosUltimaHora) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' - Total de fluxos suspensos por erro 
 Q.Clear() 
 Q.Add("SELECT COUNT(1) TOTAL FROM Z_WFMODELOINSTANCIAS WHERE SITUACAO = 5") 
 Q.Active = True 
 qtdeFluxosSuspensos = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Fluxos suspensos por erro....................: " + CStr(qtdeFluxosSuspensos) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' -  Total de fluxos suspensos por erro na última hora 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("SELECT COUNT(1) TOTAL FROM Z_WFMODELOINSTANCIAS WHERE SITUACAO = 5 AND INICIO >= (SYSDATE-1/24)") 
 Else 
    Q.Add("SELECT COUNT(1) TOTAL FROM Z_WFMODELOINSTANCIAS WHERE SITUACAO = 5 AND INICIO >= DATEADD(hour,-1,GETDATE())") 
 End If 
 Q.Active = True 
 qtdeFluxosSuspensosUltimaHora = Q.FieldByName("TOTAL").AsInteger 
 Q.Active = False 
 emailBody = emailBody + Chr(13) + Chr(10) + "Fluxos suspensos por erro na ultima hora.....: " + CStr(qtdeFluxosSuspensosUltimaHora) + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) 
 
 ' Modelos (nome e versão) com maiores tamanhos em bytes na base de dados (persistidos). 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("Select length(mi.conteudo) tamanho, mi.Handle,md.Nome,md.versao from z_wfmodeloinstancias mi,z_wfmodelos md where md.Handle = mi.modelo And ROWNUM <= 20 and mi.conteudo is not null order by tamanho desc") 
 Else 
    Q.Add("select top 20 datalength(mi.conteudo) tamanho, mi.handle, md.nome, md.versao from z_wfmodeloinstancias mi, z_wfmodelos md where md.Handle = mi.modelo order by tamanho desc") 
 End If 
 Q.Active = True 
 emailBody = emailBody + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Modelos (nome e versão) com maiores tamanhos em bytes na base de dados (persistidos):" + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) + Chr(13) + Chr(10) 
 While (Not Q.EOF) 
   emailBody = emailBody + "Tamanho: " + Q.FieldByName("tamanho").AsString + "   Instancia: " + Q.FieldByName("handle").AsString + "    Modelo: " + Q.FieldByName("nome").AsString  + " v." + Q.FieldByName("versao").AsString 
   emailBody = emailBody + Chr(13) + Chr(10) 
   Q.Next 
 Wend 
 Q.Active = False 
 
 ' Modelo (nome e versão) dos fluxos mais suspensos por erro na última hora. 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("SELECT COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MD.Handle = MI.MODELO And SITUACAO = 5 And INICIO >= (SYSDATE-1/24) GROUP BY MD.Nome, MD.VERSAO order by TOTAL desc") 
 Else 
	Q.Add("SELECT COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MD.Handle = MI.MODELO And SITUACAO = 5 And INICIO >= DateAdd(hour,-1,GETDATE()) GROUP BY MD.Nome, MD.VERSAO order by TOTAL desc") 
 End If 
 Q.Active = True 
 emailBody = emailBody + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Modelo (nome e versão) dos fluxos mais suspensos por erro na última hora." + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) + Chr(13) + Chr(10) 
 While (Not Q.EOF) 
   emailBody = emailBody + "Total: " + Q.FieldByName("TOTAL").AsString + "    Modelo: " + Q.FieldByName("FLUXO").AsString  + " v." + Q.FieldByName("VERSAO").AsString 
   emailBody = emailBody + Chr(13) + Chr(10) 
   Q.Next 
 Wend 
 Q.Active = False 
 
 
 ' Modelo (nome e versão) dos fluxos mais suspensos por erro nos últimos 30 dias. 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
    Q.Add("SELECT COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MD.Handle = MI.MODELO And SITUACAO = 5 And INICIO >= add_months(SYSDATE, -30) GROUP BY MD.Nome, MD.VERSAO order by total desc") 
 Else 
    Q.Add("SELECT COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MD.Handle = MI.MODELO And SITUACAO = 5 And INICIO >= DateAdd(Month,-30,GETDATE()) GROUP BY MD.Nome, MD.VERSAO order by total desc") 
 End If 
 Q.Active = True 
 emailBody = emailBody + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Modelo (nome e versão) dos fluxos mais suspensos por erro nos últimos 30 dias."  + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) + Chr(13) + Chr(10) 
 While (Not Q.EOF) 
   emailBody = emailBody + "Total: " + Q.FieldByName("TOTAL").AsString + "    Modelo: " + Q.FieldByName("FLUXO").AsString  + " v." + Q.FieldByName("VERSAO").AsString 
   emailBody = emailBody + Chr(13) + Chr(10) 
   Q.Next 
 Wend 
 Q.Active = False 
 
 ' Modelo (nome e versão) dos fluxos que mais geraram atividades na última hora. (quantidade de z_wfmodeloinstanciaatividades) 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("SELECT COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAATIVIDADES AT, Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MI.Handle = AT.MODELOINSTANCIA And MD.Handle = MI.MODELO And AT.INICIO >= (SYSDATE-1/24) and rownum <= 10 GROUP BY MD.Nome, MD.VERSAO ORDER BY TOTAL desc") 
 Else 
    Q.Add("SELECT TOP 10 COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAATIVIDADES AT, Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MI.Handle = AT.MODELOINSTANCIA And MD.Handle = MI.MODELO And AT.INICIO >= DateAdd(hour,-1,GETDATE()) GROUP BY MD.Nome, MD.VERSAO ORDER BY TOTAL desc") 
 End If 
 Q.Active = True 
 emailBody = emailBody + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Modelo (nome e versão) dos fluxos que mais geraram atividades na última hora. (quantidade de z_wfmodeloinstanciaatividades)" + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) + Chr(13) + Chr(10) 
 While (Not Q.EOF) 
   emailBody = emailBody + "Total: " + Q.FieldByName("TOTAL").AsString + "    Modelo: " + Q.FieldByName("FLUXO").AsString  + " v." + Q.FieldByName("VERSAO").AsString 
   emailBody = emailBody + Chr(13) + Chr(10) 
   Q.Next 
 Wend 
 Q.Active = False 
 
 ' Modelo (nome e versão) dos fluxos que mais foram instanciados na última hora. 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("SELECT COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MD.HANDLE = MI.MODELO AND MI.INICIO >= (SYSDATE-1/24) and rownum <= 10 GROUP BY MD.NOME, MD.VERSAO order by total desc") 
 Else 
    Q.Add("SELECT TOP 10 COUNT(1) TOTAL, MD.NOME FLUXO, MD.VERSAO FROM Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS MD WHERE MD.HANDLE = MI.MODELO AND MI.INICIO >= DATEADD(hour, -1, GETDATE()) GROUP BY MD.NOME, MD.VERSAO order by total desc") 
 End If 
 Q.Active = True 
 emailBody = emailBody + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Modelo (nome e versão) dos fluxos que mais foram instanciados na última hora." + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) + Chr(13) + Chr(10) 
 While (Not Q.EOF) 
   emailBody = emailBody + "Total: " + Q.FieldByName("TOTAL").AsString + "    Modelo: " + Q.FieldByName("FLUXO").AsString  + " v." + Q.FieldByName("VERSAO").AsString 
   emailBody = emailBody + Chr(13) + Chr(10) 
   Q.Next 
 Wend 
 Q.Active = False 
 
 ' Macros que mais levaram tempo para concluir o processamento na última hora. 
 Q.Clear() 
 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
 	Q.Add("SELECT (cast(msg.concluidaem as timestamp) - cast(msg.criadaem as timestamp)) duracao, MSG.HANDLE, MOD.NOME FLUXO, MOD.VERSAO, AT.TITULO ATIVIDADE, MSG.MODELOINSTANCIA FROM Z_WFMENSAGENS MSG,  Z_WFMODELOINSTANCIAATIVIDADES AT, Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS Mod WHERE MSG.GUID In (Select ASSUNTO FROM Z_FILAMENSAGENS WHERE FILA = (Select HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_VBAPROCESS$') AND ESTADO = 3 And DATAHORA > (SYSDATE-1/24)) And AT.Handle = MSG.ATIVIDADE And MI.Handle = AT.MODELOINSTANCIA And Mod.Handle  = MI.MODELO and rownum <= 20 ORDER BY DURACAO DESC") 
 Else 
    Q.Add("SELECT TOP 20 (MSG.CONCLUIDAEM - MSG.CRIADAEM) DURACAO, MSG.HANDLE, MOD.NOME FLUXO, MOD.VERSAO, AT.TITULO ATIVIDADE, MSG.MODELOINSTANCIA FROM Z_WFMENSAGENS MSG,  Z_WFMODELOINSTANCIAATIVIDADES AT, Z_WFMODELOINSTANCIAS MI, Z_WFMODELOS Mod WHERE MSG.GUID In (Select ASSUNTO FROM Z_FILAMENSAGENS WHERE FILA = (Select HANDLE FROM Z_FILAS WHERE NOME = '$BENNER_WORKFLOW_VBAPROCESS$') AND ESTADO = 3 And DATAHORA > DateAdd(hour,-1, GETDATE())) And AT.Handle = MSG.ATIVIDADE And MI.Handle = AT.MODELOINSTANCIA And Mod.Handle  = MI.MODELO ORDER BY DURACAO DESC") 
 End If 
 Q.Active = True 
 emailBody = emailBody + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Macros que mais levaram tempo para concluir o processamento na última hora." + "               custo do sql: " + FormatDateTime2("hh:nn:ss", ServerNow - inicio) + Chr(13) + Chr(10) 
 While (Not Q.EOF) 
	 If SQLServer = "ORACLE8I" Or SQLServer = "ORACLE9I" Then 
	 	 emailBody = emailBody + "Handle: " + Q.FieldByName("HANDLE").AsString + "    Instancia: " + Q.FieldByName("MODELOINSTANCIA").AsString +  "     Duracao: " + Q.FieldByName("DURACAO").AsString + "    Modelo: " + Q.FieldByName("FLUXO").AsString  + " v." + Q.FieldByName("VERSAO").AsString + "   Atividade: " + Q.FieldByName("ATIVIDADE").AsString 
	 Else 
    emailBody = emailBody + "Handle: " + Q.FieldByName("HANDLE").AsString + "    Instancia: " + Q.FieldByName("MODELOINSTANCIA").AsString + "     Duracao: " + FormatDateTime2("hh:nn:ss", Q.FieldByName("DURACAO").AsDateTime) + "    Modelo: " + Q.FieldByName("FLUXO").AsString  + " v." + Q.FieldByName("VERSAO").AsString + "   Atividade: " + Q.FieldByName("ATIVIDADE").AsString 
  End If 
  emailBody = emailBody + Chr(13) + Chr(10) 
  Q.Next 
 Wend 
 Q.Active = False 
 
 Set Q = Nothing 
 
 Dim objMail As Object 
 Set objMail = NewMail 
 
 objMail.Clear 
 objMail.From = emailRemetente 
 objMail.Subject = "Relatório de monitoramento do Workflow - " + CStr(ServerNow) 
 objMail.Priority = 0 
 objMail.SendTo = emailDestinatarios 
 objMail.Text.Add(emailBody) 
 objMail.Send 
 
 Set objMail = Nothing 
End Sub 
