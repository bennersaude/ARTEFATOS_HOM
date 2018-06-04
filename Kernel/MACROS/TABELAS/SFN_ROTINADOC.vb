'HASH: AB0789FBA6AD8041F9809C39AB65E88E
'Macro: SFN_ROTINADOC
'#Uses "*bsShowMessage"

Option Explicit

Public Sub BOTAOAGENDAR_OnClick()
  Dim qr As Object
  Dim qr1 As Object
  Dim vSituacao As Integer
  Dim vTabela As String
  Dim vLegendaAgendamento As Integer
  Dim VLegendaAberta As Integer
  Dim VLegendaProcessada As Integer
  Set qr = NewQuery
  Set qr1 = NewQuery
  vTabela = "SFN_ROTINADOC"
  vLegendaAgendamento = 3
  VLegendaAberta = 1
  VLegendaProcessada = 5
  qr.Clear
  qr.Add("SELECT SITUACAO FROM " + vTabela + " WHERE HANDLE = :pHANDLE")
  qr.ParamByName("pHandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qr.Active = True
  vSituacao = qr.FieldByName("SITUACAO").AsInteger
  If vSituacao <> vLegendaAgendamento Then
    If vSituacao = VLegendaAberta Then
      If bsShowMessage("Confirme o agendamento da rotina", "Q") = vbYes Then '(6=yes, 7=não)
        qr1.Clear
		If Not InTransaction Then StartTransaction
	        qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
        	qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        	qr1.ParamByName("pSituacao").AsInteger = vLegendaAgendamento
        	qr1.ExecSQL
        If InTransaction Then Commit
      End If
    Else
      bsShowMessage("Rotina já foi processada.", "I")
    End If
  Else
    If bsShowMessage("Rotina já está agendada. Para retirar o agendamento pressione 'SIM'", "Q") = vbYes Then
      qr1.Clear
 	  If Not InTransaction Then StartTransaction
	      qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
      	qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      	If (CurrentQuery.FieldByName("PROCESSADODATA").IsNull) Then
	        qr1.ParamByName("pSituacao").AsInteger = VLegendaAberta
      	Else
	        qr1.ParamByName("pSituacao").AsInteger = VLegendaProcessada
      	End If
      	qr1.ExecSQL
      If InTransaction Then Commit
    End If
  End If
  Set qr = Nothing
  Set qr1 = Nothing
  SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
End Sub

Public Sub BOTAOALTERARTITULO_OnClick()


	Dim QUpdate As Object
	Dim Query As Object
	Dim Obj As Object
	Dim Resultado As Integer


	Set Query = NewQuery

	Set QUpdate = NewQuery
    Set Obj = CreateBennerObject("CS.Formularios")

    Dim vCondicao As String

	Query.Add("SELECT HANDLE FROM Z_FORMULARIOS WHERE NOME = 'ALTERATITULO'")
	Query.Active = True

    Obj.CreateObjForm("Alteração do Tipo do cobrança")

	If Query.FieldByName("HANDLE").AsInteger = 0 Then   		' SMS 95929 - Paulo Melo - 18/04/2008
		bsShowMessage("Formulário não encontrado", "E")
		Exit Sub
	End If

    Resultado = Obj.Execute(Query.FieldByName("HANDLE").AsInteger)


	If Resultado > 0 Then
		If Obj.Fields("TABGERACAO") = 1 Then
			If bsShowMessage("Confirma a alteração dos títulos desta rotina para Conta-corrente ?","Q") = vbYes Then
				While Not Query.EOF

					If Obj.Fields("CONDICAO") = 1 Then

					  If Not InTransaction Then StartTransaction
					    QUpdate.Add("UPDATE SFN_DOCUMENTO SET TABGERACAO = 1 ")
					    QUpdate.Add("WHERE ROTINADOC = :PROTINADOC ")
						Select Case Obj.Fields("FILTRO")
							Case "1"
								QUpdate.Add(" AND VALOR > :PVALOR")

							Case "2"
								QUpdate.Add(" AND VALOR < :PVALOR")

							Case "3"
								QUpdate.Add(" AND VALOR = :PVALOR")
						End Select

					    QUpdate.ParamByName("PVALOR").AsFloat = CCur(Obj.Fields("VALOR"))
					    QUpdate.ParamByName("PROTINADOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
					    QUpdate.ExecSQL
					  If InTransaction Then Commit

					Else
					  If Not InTransaction Then StartTransaction
					    QUpdate.Add("UPDATE SFN_DOCUMENTO SET TABGERACAO = 1 ")
					    QUpdate.Add("WHERE ROTINADOC = :PROTINADOC")
					    QUpdate.ParamByName("PROTINADOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
					    QUpdate.ExecSQL
					  If InTransaction Then Commit

					End If


				    Query.Next
				Wend
			End If

		Else

			If bsShowMessage("Confirma a alteração dos títulos desta rotina para Boleto ?", "Q") = vbYes Then

				While Not Query.EOF

					If Obj.Fields("CONDICAO") = 1 Then
					  If Not InTransaction Then StartTransaction
					    QUpdate.Add("UPDATE SFN_DOCUMENTO SET TABGERACAO = 3 ")
					    QUpdate.Add("                         ,INSTRUCAO1 = '"+ Obj.Fields("INSTRUCAO1")+"'")
					    QUpdate.Add("                         ,INSTRUCAO2 = '"+ Obj.Fields("INSTRUCAO2")+"'")
					    QUpdate.Add("                         ,INSTRUCAO3 = '"+ Obj.Fields("INSTRUCAO3")+"'")
					    QUpdate.Add("                         ,INSTRUCAO4 = '"+ Obj.Fields("INSTRUCAO4")+"'")
					    QUpdate.Add("                         ,INSTRUCAO5 = '"+ Obj.Fields("INSTRUCAO5")+"'")
					    QUpdate.Add("WHERE ROTINADOC = :PROTINADOC")

						Select Case Obj.Fields("FILTRO")
							Case "1"
								QUpdate.Add(" AND VALOR >:PVALOR")

							Case "2"
								QUpdate.Add(" AND VALOR <:PVALOR")

							Case "3"
								QUpdate.Add(" AND VALOR =:PVALOR")
						End Select

					    QUpdate.ParamByName("PVALOR").AsFloat = CCur(Obj.Fields("VALOR"))
					    QUpdate.ParamByName("PROTINADOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
					    QUpdate.ExecSQL
					 If InTransaction Then Commit

					Else
					  If Not InTransaction Then StartTransaction
					    QUpdate.Add("UPDATE SFN_DOCUMENTO SET TABGERACAO = 3 ")
					    QUpdate.Add("                         ,INSTRUCAO1 = '"+ Obj.Fields("INSTRUCAO1")+"'")
					    QUpdate.Add("                         ,INSTRUCAO2 = '"+ Obj.Fields("INSTRUCAO2")+"'")
					    QUpdate.Add("                         ,INSTRUCAO3 = '"+ Obj.Fields("INSTRUCAO3")+"'")
					    QUpdate.Add("                         ,INSTRUCAO4 = '"+ Obj.Fields("INSTRUCAO4")+"'")
					    QUpdate.Add("                         ,INSTRUCAO5 = '"+ Obj.Fields("INSTRUCAO5")+"'")
					    QUpdate.Add("WHERE ROTINADOC = :PROTINADOC")
					    QUpdate.ParamByName("PROTINADOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
					    QUpdate.ExecSQL
					  If InTransaction Then Commit

					End If

				    Query.Next
				Wend

			End If
		End If
	End If

	Set Obj = Nothing
	Set Query = Nothing
	Set QUpdate = Nothing

End Sub


Public Sub BOTAOCANCELARDOCUMENTOS_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT D.HANDLE ")
  SQL.Add("  FROM SFN_DOCUMENTO D ")
  SQL.Add("  Join SFN_NOTA_DOCUMENTO ND On ND.DOCUMENTO = D.Handle")
  SQL.Add("  Join SFN_NOTA N On N.Handle = ND.NOTA ")
  SQL.Add(" WHERE D.ROTINADOC = :HROTINA")
  SQL.Add("   And N.TABORIGEM = 1	")
  SQL.Add("   AND D.CANCDATA IS NOT NULL	")
  SQL.Add("   AND D.BAIXADATA IS NULL	")
  SQL.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    If bsShowMessage("Exitem notas fiscais vinculadas e conciliadas a documentos." + Chr(13) + "Deseja realmente cancelar os documentos?", "Q") = vbNo Then
	  SQL.Active = False
	  Set SQL = Nothing
	  Exit Sub
	End If
  End If

  SQL.Active = False
  Set SQL = Nothing

  Dim Obj As Object

  Set Obj = CreateBennerObject("SFNCancel.Cancelamento")
  Obj.Inicializar(CurrentSystem)
  Obj.CancelaDocRot(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Obj.Finalizar
  Set Obj = Nothing

  WriteAudit("I", HandleOfTable("SFN_ROTINADOC"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina Documento - Cancelamento de Documentos")
End Sub

Public Sub BOTAOCONFIRMARINTEGRACAO_OnClick()
  Dim SfnRotinaDocBLL As CSBusinessComponent
  Set SfnRotinaDocBLL = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnRotinaDocBLL, Benner.Saude.Financeiro.Business")

  SfnRotinaDocBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  SfnRotinaDocBLL.Execute("ConfirmarIntegracao")

  bsShowMessage("Confirmação da Integração concluída!", "I")
End Sub

Public Sub BOTAODUPLICAR_OnClick()

  If CurrentQuery.State <> 1 Then
		bsShowMessage("Os parâmetros não podem estar em edição", "I")
		Exit Sub
  End If

  Dim INTERFACE0002 As Object
  Dim vsMensagem As String


  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

	INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0142", _
					   "Duplicar rotina documento",  _
					   0, _
					   250, _
					   430, _
					   False, _
					   vsMensagem, _
					   Null)


  Set INTERFACE0002 = Nothing

  WriteAudit("D", HandleOfTable("SFN_ROTINAFINFAT"), CurrentQuery.FieldByName("HANDLE").AsInteger,"Rotina de Documento - Duplicação")

  RefreshNodesWithTable("SFN_ROTINADOC")

End Sub

Public Sub BOTAOIMPRIMIRNOTA_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If

  Dim Obj As Object

  Set Obj = CreateBennerObject("SamImpressao.NotaFiscal")
  Obj.Inicializar(CurrentSystem)
  Obj.ImprimirRotina(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Obj.Finalizar
  Set Obj = Nothing

  WriteAudit("I", HandleOfTable("SFN_ROTINADOC"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina Documento - Impressão de Notas")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If CurrentQuery.FieldByName("TABFILTRO").AsInteger = 3 Then 'várias rotinas
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT COUNT(*) NREC FROM SFN_ROTINADOC_ROTFIN WHERE ROTINADOC=:ROTINADOC")
    SQL.ParamByName("ROTINADOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If SQL.FieldByName("NREC").AsInteger = 0 Then
      bsShowMessage("Devem existir registros de rotinas financeiras.", "I")
      Set SQL = Nothing
      Exit Sub
    End If
    Set SQL = Nothing
  End If

  Dim CanContinue As Boolean
  VerificaSeProcessada(CanContinue)
  If Not CanContinue Then
    Exit Sub
  End If

  Dim Obj As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String
  Dim vcContainer As CSDContainer

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If


  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0041.RotinaDocumento")
	Obj.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  Else
	Set vcContainer = NewContainer

    vcContainer.AddFields("HANDLE:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "Financeiro", _
                       			   	 "RotinaDocumento_ProcessaDocumento", _
                     			     "Rotina Documento - Processamento", _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINADOC", _
                                     "SITUACAO", _
                                     "", _
                                     "", _
                                     "P", _
                                     False, _
                                     vsMensagem, _
                                     vcContainer)

	If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
    End If
  End If
  Set Obj = Nothing
  Set vcContainer = Nothing

End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If


  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT D.HANDLE ")
  SQL.Add("  FROM SFN_DOCUMENTO D ")
  SQL.Add("  Join SFN_NOTA_DOCUMENTO ND On ND.DOCUMENTO = D.Handle")
  SQL.Add("  Join SFN_NOTA N On N.Handle = ND.NOTA ")
  SQL.Add(" WHERE D.ROTINADOC = :HROTINA")
  SQL.Add("   And N.TABORIGEM = 1	")
  SQL.ParamByName("HROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    If bsShowMessage("Exitem notas fiscais vinculadas e conciliadas a documentos." + Chr(13) + "Deseja realmente cancelar a rotina?", "Q") = vbNo Then
	  SQL.Active = False
	  Set SQL = Nothing
	  Exit Sub
	End If
  Else
    If bsShowMessage("Cancelar a execução da rotina documento ?", "Q") = vbNo Then
      SQL.Active = False
	  Set SQL = Nothing
      Exit Sub
    End If
  End If

  SQL.Active = False
  Set SQL = Nothing

  If VisibleMode Then
    If CurrentQuery.FieldByName("CANCELADODATA").IsNull Then
       Set Obj = CreateBennerObject("BSINTERFACE0041.RotinaDocumento")
       Obj.Cancelar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
       CurrentQuery.Active = False
       CurrentQuery.Active = True
    Else
       bsShowMessage("A rotina já foi cancelada.", "I")
    End If
  Else
     Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
     viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                   "Financeiro", _
                                   "RotinaDocumento_CancelaDocumento", _
                                   "Rotina Documento - Cancelamento", _
                                   CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                   "SFN_ROTINADOC", _
                                   "SITUACAO", _
                                   "", _
                                   "", _
                                   "C", _
                                   False, _
                                   vsMensagem, _
                                   Null)

     If viRetorno = 0 Then
       bsShowMessage("Processo enviado para execução no servidor!", "I")
     Else
       bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
     End If
 End If
 Set Obj = Nothing

End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  Dim Obj As Object
  Dim sql As Object
  Dim Aux As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If
  If VisibleMode Then
	  Set Obj = CreateBennerObject("SamImpressao.Boleto")
	  Obj.Inicializar(CurrentSystem)
	  Obj.ImprimirRotina(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
	  Obj.Finalizar
	  Set Obj = Nothing

	  WriteAudit("I", HandleOfTable("SFN_ROTINADOC"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina Documento - Impressão")
  Else
    Set Aux = NewQuery
    Aux.Active = False
    Aux.Clear
    Aux.Add("SELECT MIN(NUMERO) INICIAL, MAX(NUMERO) FINAL")
    Aux.Add("FROM SFN_DOCUMENTO")
    Aux.Add("WHERE ROTINADOC = :HANDLE")
    Aux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    Aux.Active = True

    UserVar("HDocInicial") = Aux.FieldByName("INICIAL").AsString
    UserVar("HDocFinal") = Aux.FieldByName("FINAL").AsString

    Set sql = NewQuery
    sql.Active = False
	sql.Clear
	sql.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'SFN-DC-998'")
	sql.Active = True

	If sql.EOF Then
		bsShowMessage("Relatório 'SFN-DC-998' não encontrado !" + Chr(10) + _
									"Verique se: " + Chr(10) + _
									"- o relatório foi importado;" + Chr(10) + _
									"- o código do relatório está correto." + Chr(10) + _
									Chr(10) + _
									"Para importar ou corrigir o código entre no" + Chr(10) + _
  									"Módulo ´Adm´ / carga: ´Gerador de relatórios/Relatórios/...", "I")
		Exit Sub
	Else
		Dim rel As CSReportPrinter
		Set rel = NewReport(sql.FieldByName("HANDLE").AsInteger)
		rel.Preview

		Set rel = Nothing
		'ReportPreview(sql.FieldByName("HANDLE").Value, "", True, False)
		'RefreshNodesWithTable("SFN_ROTINADOC")
	End If
	Set sql = Nothing
	Set Aux = Nothing
  End If

End Sub

Public Sub BOTAOIMPRIMIRARQUIVO_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If

  Set Obj = CreateBennerObject("SAMBOLETO.Exportar")
  Obj.RotinaDoc(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Obj = Nothing

  WriteAudit("I", HandleOfTable("SFN_ROTINADOC"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Rotina Documento - Impressão em Arquivo")

End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").Value <> "1" Then
    CanContinue = False
    bsShowMessage("A Rotina já foi processada.", "E")
    Exit Sub
  Else
    CanContinue = True
  End If
End Sub

Public Sub DATA_OnExit()

If CurrentQuery.State = 2 Or CurrentQuery.State = 3 Then  'SMS 85141 - Paulo Melo - 31/07/2007

  If CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull Then
    CurrentQuery.FieldByName("VENCIMENTOINICIAL").Value = CurrentQuery.FieldByName("DATA").AsDateTime
  End If

  If CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull Then
    CurrentQuery.FieldByName("VENCIMENTOFINAL").Value = CurrentQuery.FieldByName("DATA").AsDateTime
  End If
End If
End Sub

Public Sub TABFILTRO_OnChange()
  If TABFILTRO.PageIndex <2 Then
    Dim sql As Object
    Set sql = NewQuery
    sql.Add("SELECT COUNT(*) NREC FROM SFN_ROTINADOC_ROTFIN WHERE ROTINADOC=:ROTINADOC")
    sql.ParamByName("ROTINADOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sql.Active = True
    If sql.FieldByName("NREC").AsInteger >0 Then
      CurrentQuery.FieldByName("TABFILTRO").AsInteger = 3
      bsShowMessage("Existem registro de rotinas financeiras abaixo desta rotina documento. Devem ser excluídas para ser alterado o filtro.", "I")
    End If
    Set sql = Nothing
  End If
End Sub

Public Sub TABFILTRO_OnChanging(AllowChange As Boolean)
  VerificaSeProcessada(AllowChange)
End Sub

Public Sub TABLE_AfterScroll()

  If VisibleMode Then
  	If NodeInternalCode = 600 Then
     	'TABTIPOROTINA.Pages(0).Visible = False
     	'TABTIPOROTINA.Pages(1).Visible = True

     	BOTAOAGENDAR.Visible = False
     	BOTAOCANCELAR.Visible = False
     	BOTAOIMPRIMIR.Visible = False
     	BOTAOIMPRIMIRARQUIVO.Visible = False
     	BOTAOIMPRIMIRNOTA.Visible = False
     	BOTAOPROCESSAR.Visible = False
     	BOTAOALTERARTITULO.Visible = False

   	Else
'    	TABTIPOROTINA.Pages(0).Visible = True
'     	TABTIPOROTINA.Pages(1).Visible = False

     	BOTAOAGENDAR.Visible = True
     	BOTAOCANCELAR.Visible = True
     	BOTAOIMPRIMIR.Visible = True
     	BOTAOIMPRIMIRARQUIVO.Visible = True
     	BOTAOIMPRIMIRNOTA.Visible = True
     	BOTAOPROCESSAR.Visible = True
     	BOTAOALTERARTITULO.Visible = True

   	End If

   	Dim SQLParamIntegracao As BPesquisa
   	Set SQLParamIntegracao = NewQuery

	SQLParamIntegracao.Add("SELECT 1")
	SQLParamIntegracao.Add("FROM ADM_PARAMINTEGRACAOCORPBENNER")
	SQLParamIntegracao.Active = True

	If SQLParamIntegracao.EOF Then
		BOTAOCONFIRMARINTEGRACAO.Visible = False
	Else
		Dim SQLTipoDocumento As BPesquisa
		Set SQLTipoDocumento = NewQuery

		SQLTipoDocumento.Add("SELECT TABINTEGRACAOCORPORATIVO")
		SQLTipoDocumento.Add("FROM SFN_TIPODOCUMENTO")
		SQLTipoDocumento.Add("WHERE HANDLE = :HTIPODOCUMENTO")
		SQLTipoDocumento.ParamByName("HTIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
		SQLTipoDocumento.Active = True
		If SQLTipoDocumento.FieldByName("TABINTEGRACAOCORPORATIVO").AsInteger = 2 Then
			BOTAOCONFIRMARINTEGRACAO.Visible = True
		Else
			BOTAOCONFIRMARINTEGRACAO.Visible = False
		End If
		Set SQLTipoDocumento = Nothing
	End If

   	Set SQLParamIntegracao = Nothing
  ElseIf WebMode Then
    BOTAOAGENDAR.Visible = True
    BOTAOCANCELAR.Visible = True
    BOTAOIMPRIMIR.Visible = True
    BOTAOIMPRIMIRARQUIVO.Visible = True
    BOTAOIMPRIMIRNOTA.Visible = True
    BOTAOPROCESSAR.Visible = True
    BOTAOALTERARTITULO.Visible = True
 End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If VisibleMode Then
    If NodeInternalCode <> 600 Then

      If Not CurrentQuery.FieldByName("TIPOFATURAMENTO").IsNull Then
        If CurrentQuery.FieldByName("COMPETFIN").IsNull Or  _
          CurrentQuery.FieldByName("ROTINAFIN").IsNull Then
          If TABFILTRO.PageIndex = 1 Then
            bsShowMessage("Se informar o tipo de faturamento deve informar a competência e a rotina financeira.", "E")
            TIPOFATURAMENTO.SetFocus
            CanContinue =False
          Else
            CurrentQuery.FieldByName("TIPOFATURAMENTO").Value = Null
          End If
          Exit Sub
        End If
      End If

    'Pinheiro sms 25574
      Dim sql As Object
      Set sql=NewQuery
      sql.Add("SELECT INCLUIFATURASEMDOCUMENTO FROM SFN_TIPODOCUMENTO WHERE HANDLE = :PHANDLE")
      sql.ParamByName("PHANDLE").AsInteger=CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
      sql.Active=True
      If sql.FieldByName("INCLUIFATURASEMDOCUMENTO").AsString = "S" Then
        If (CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString = "S") Then
          bsShowMessage("O campo 'Inclui faturas atrasadas' não pode ser marcado porque no tipo de documento escolhido está marcada a opção"+ _
                 "'Inclui faturas sem documento'", "E")
          Set sql=Nothing
          CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString = "N"
          CanContinue = False
          Exit Sub
        End If
        If CurrentQuery.FieldByName("DATAATRASO").IsNull Then
          bsShowMessage("O campo 'Antes do dia' deve estar preenchido porque no tipo de documento escolhido a opção"+ _
                 "'Inclui faturas sem documento' está marcada", "E")
          Set sql=Nothing
          CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString = "N"
          CanContinue = False
          Exit Sub
        End If
      End If
      Set sql=Nothing
      'FIM Pinheiro sms 25574

      If(CurrentQuery.FieldByName("TABFILTRO").AsInteger <> 2) Then
        CurrentQuery.FieldByName("TIPOFATURAMENTO").Clear
        CurrentQuery.FieldByName("COMPETFIN").Clear
        CurrentQuery.FieldByName("ROTINAFIN").Clear
      End If

'      If(CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString ="S") And (CurrentQuery.FieldByName("DATAATRASO").IsNull) Then
'	     bsShowMessage("A data 'Antes do dia' deve estar preenchida !", "E")
'	     CanContinue =False
'	     Exit Sub
'      End If

      If(CurrentQuery.FieldByName("TABFILTRO").AsInteger =1) Then
        If(CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString ="S") And (CurrentQuery.FieldByName("TABTIPOGERACAO").AsInteger=1)Then
          If CurrentQuery.FieldByName("DATAATRASO").IsNull Then
             bsShowMessage("A data 'Antes do dia' deve estar preenchida !", "E")
             CanContinue =False
             Exit Sub
          End If

          If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime <=CurrentQuery.FieldByName("DATAATRASO").AsDateTime Then
            bsShowMessage("A data 'Antes do dia' deve ser menor que o Vencimento inicial !", "E")
            CanContinue =False
            Exit Sub
          End If
        End If

        If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime >CurrentQuery.FieldByName("VENCIMENTOFINAL").AsDateTime Then
          bsShowMessage("A data 'Vencimento inicial' deve ser menor ou igual que a Data 'Vencimento final'!", "E")
          VENCIMENTOINICIAL.SetFocus
          CanContinue =False
          TABLE.ActivePage(0)
          Exit Sub
        End If

        If CurrentQuery.FieldByName("TABTIPOGERACAO").AsInteger = 1 Then
          If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime < CurrentQuery.FieldByName("DATA").AsDateTime Then
             bsShowMessage("A data 'Vencimento inicial' deve ser maior ou igual que a Data base de geração do documento!", "E")
             VENCIMENTOINICIAL.SetFocus
             CanContinue = False
             Exit Sub
          End If
         End If
      End If
    End If
  ElseIf WebMode Then ' WEB
  	 If Not CurrentQuery.FieldByName("TIPOFATURAMENTO").IsNull Then
       If CurrentQuery.FieldByName("COMPETFIN").IsNull Or  _
          CurrentQuery.FieldByName("ROTINAFIN").IsNull Then
          If TABFILTRO.PageIndex = 1 Then
            bsShowMessage("Se informar o tipo de faturamento deve informar a competência e a rotina financeira.", "E")
            TIPOFATURAMENTO.SetFocus
            CanContinue =False
          Else
            CurrentQuery.FieldByName("TIPOFATURAMENTO").Value = Null
          End If
          Exit Sub
       End If
     End If

    'Pinheiro sms 25574
    Dim sql2 As Object
    Set sql2=NewQuery
    sql2.Add("SELECT INCLUIFATURASEMDOCUMENTO FROM SFN_TIPODOCUMENTO WHERE HANDLE = :PHANDLE")
    sql2.ParamByName("PHANDLE").AsInteger=CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
    sql2.Active=True
    If sql2.FieldByName("INCLUIFATURASEMDOCUMENTO").AsString = "S" Then
      If (CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString = "S") Then
        bsShowMessage("O campo 'Inclui faturas atrasadas' não pode ser marcado porque no tipo de documento escolhido está marcada a opção"+ _
               "'Inclui faturas sem documento'", "E")
        Set sql2=Nothing
        CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString = "N"
        CanContinue = False
        Exit Sub
      End If
      If CurrentQuery.FieldByName("DATAATRASO").IsNull Then
        bsShowMessage("O campo 'Antes do dia' deve estar preenchido porque no tipo de documento escolhido a opção"+ _
               "'Inclui faturas sem documento' está marcada", "E")
        Set sql2=Nothing
        CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString = "N"
        CanContinue = False
        Exit Sub
      End If
    End If
    Set sql2=Nothing
    'FIM Pinheiro sms 25574

    If(CurrentQuery.FieldByName("TABFILTRO").AsInteger <> 2) Then
      CurrentQuery.FieldByName("TIPOFATURAMENTO").Clear
      CurrentQuery.FieldByName("COMPETFIN").Clear
      CurrentQuery.FieldByName("ROTINAFIN").Clear
    End If

    If(CurrentQuery.FieldByName("TABFILTRO").AsInteger =1) Then
      If(CurrentQuery.FieldByName("INCLUIFATURASATRASADAS").AsString ="S") And (CurrentQuery.FieldByName("TABTIPOGERACAO").AsInteger=1)Then
        If CurrentQuery.FieldByName("DATAATRASO").IsNull Then
           bsShowMessage("A data 'Antes do dia' deve estar preenchida !", "E")
           CanContinue =False
           Exit Sub
        End If

        If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime <=CurrentQuery.FieldByName("DATAATRASO").AsDateTime Then
           bsShowMessage("A data 'Antes do dia' deve ser menor que o Vencimento inicial !", "E")
           CanContinue =False
           Exit Sub
        End If
      End If

      If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime >CurrentQuery.FieldByName("VENCIMENTOFINAL").AsDateTime Then
        bsShowMessage("A data 'Vencimento inicial' deve ser menor ou igual que a Data 'Vencimento final'!", "E")
        VENCIMENTOINICIAL.SetFocus
        CanContinue =False
        TABLE.ActivePage(0)
        Exit Sub
      End If

      If CurrentQuery.FieldByName("TABTIPOGERACAO").AsInteger = 1 Then
        If CurrentQuery.FieldByName("VENCIMENTOINICIAL").AsDateTime < CurrentQuery.FieldByName("DATA").AsDateTime Then
           bsShowMessage("A data 'Vencimento inicial' deve ser maior ou igual que a Data base de geração do documento!", "E")
           VENCIMENTOINICIAL.SetFocus
           CanContinue = False
           Exit Sub
        End If
      End If
    End If
  End If

End Sub

Public Sub TABLE_NewRecord()
  If VisibleMode Then
  	If NodeInternalCode = 600 Then

    	'TABTIPOROTINA.Pages(0).Visible = False
     	'TABTIPOROTINA.Pages(1).Visible = True
     	'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 2

	    BOTAOAGENDAR.Visible = False
     	BOTAOCANCELAR.Visible = False
     	BOTAOIMPRIMIR.Visible = False
     	BOTAOIMPRIMIRARQUIVO.Visible = False
     	BOTAOIMPRIMIRNOTA.Visible = False
     	BOTAOPROCESSAR.Visible = False
     	BOTAOALTERARTITULO.Visible = False

   	Else
     	'TABTIPOROTINA.Pages(0).Visible = True
     	'TABTIPOROTINA.Pages(1).Visible = False
     	'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1

     	BOTAOAGENDAR.Visible = True
     	BOTAOCANCELAR.Visible = True
     	BOTAOIMPRIMIR.Visible = True
     	BOTAOIMPRIMIRARQUIVO.Visible = True
     	BOTAOIMPRIMIRNOTA.Visible = True
     	BOTAOPROCESSAR.Visible = True
     	BOTAOALTERARTITULO.Visible = True

   	End If
  ElseIf WebMode Then
    'CurrentQuery.FieldByName("TABTIPOROTINA").Value = 1

    BOTAOAGENDAR.Visible = True
    BOTAOCANCELAR.Visible = True
    BOTAOIMPRIMIR.Visible = True
    BOTAOIMPRIMIRARQUIVO.Visible = True
    BOTAOIMPRIMIRNOTA.Visible = True
    BOTAOPROCESSAR.Visible = True
    BOTAOALTERARTITULO.Visible = True
  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOAGENDAR"
			BOTAOAGENDAR_OnClick
		Case "BOTAOALTERARTITULO"
			BOTAOALTERARTITULO_OnClick
		Case "BOTAOCANCELARDOCUMENTOS"
			BOTAOCANCELARDOCUMENTOS_OnClick
		Case "BOTAOIMPRIMIRNOTA"
			BOTAOIMPRIMIRNOTA_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOIMPRIMIR"
			BOTAOIMPRIMIR_OnClick
		Case "BOTAOIMPRIMIRARQUIVO"
			BOTAOIMPRIMIRARQUIVO_OnClick
	End Select
End Sub
