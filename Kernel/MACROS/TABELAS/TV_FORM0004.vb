'HASH: A72827DD596B9F582F5EF90743666546
				'#Uses "*bsShowMessage

		Option Explicit

		Dim vs_PRECOTOTAL As String
		Dim vs_VALORPF As String
		Dim VsMensagem As String


		Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)

		  Dim procuraDll As Object
		  Dim handle As Long

		  Set procuraDll = CreateBennerObject("BSINTERFACE0005.ConsultaBeneficiario")

		  ShowPopup = False

		  handle = procuraDll.Filtro(CurrentSystem, _
							    	 1, _
							    	 BENEFICIARIO.Text)

		  If (handle <> 0) Then
		    CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = handle
		  End If

		  Dim qPlano As Object
		  Set qPlano = NewQuery

		  Dim v_Qtd As Integer

			qPlano.Add("SELECT CM.PLANO")
			qPlano.Add("  FROM SAM_BENEFICIARIO_MOD BM")
			qPlano.Add("  JOIN SAM_CONTRATO_MOD CM ON (CM.HANDLE = BM.MODULO)")
			qPlano.Add(" WHERE BM.BENEFICIARIO =:PHANDLEBENEFICIARIO AND CM.OBRIGATORIO = 'S'")
			qPlano.Add("   AND (BM.DATACANCELAMENTO IS NULL OR (BM.DATACANCELAMENTO IS NOT NULL AND BM.DATACANCELAMENTO >= :PDATAATUAL))")
			qPlano.ParamByName("PHANDLEBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
			qPlano.ParamByName("PDATAATUAL").AsDateTime = CurrentQuery.FieldByName("DATA").AsDateTime
			qPlano.Active = True

			v_Qtd = 0

			While Not qPlano.EOF

				v_Qtd = v_Qtd + 1

				qPlano.Next
			Wend

			If v_Qtd = 1 Then
				CurrentQuery.FieldByName("PLANO").AsInteger = qPlano.FieldByName("PLANO").AsInteger
			End If

			Set qPlano = Nothing

		  Set procuraDll = Nothing
		End Sub

		Public Sub BOTAOIMPRIMIR_OnClick()
		  Dim Relatorio As Object
		  Dim P_DATA As Date
		  Dim P_GRAU As Long
		  Dim P_EVENTO As Long
		  Dim P_XTHM As Long
		  Dim P_CODPAGAMENTO As Long
		  Dim P_QUANTIDADE As Long
		  Dim P_BENEFICIARIO As Long
		  Dim P_LOCALATENDIMENTO As Long
		  Dim P_CONDICAOATENDIMENTO As Long
		  Dim P_REGIMEATENDIMENTO As Long
		  Dim P_TIPOTRATAMENTO As Long
		  Dim P_OBJETIVOTRATAMENTO As Long
		  Dim P_FINALIDADETRATAMENTO As Long
		  Dim P_LOCALDEEXECUCAO As Long
		  Dim P_FINALIDADEATENDIMENTO As Long
		  Dim P_CBOS As Long
		  Dim P_RECEBEDOR As Long
		  Dim vsMensagemErro As String
		  Dim bs As CSBusinessComponent
		  Dim res As String
		  Dim P_TABXTHM As Long
		  Dim P_CONVENIO As Long
		  Dim P_FILIAL As Long
		  Dim P_ESTADO As Long
		  Dim P_MUNICIPIO As Long
		  Dim P_PLANO As Long
		  Dim P_EXECUTOR As Long
		  Dim P_ACOMODACAO As Long
		  Dim P_TECNICACIRURGICA As String


		  P_DATA = CurrentQuery.FieldByName("DATA").AsDateTime
		  P_GRAU = CurrentQuery.FieldByName("GRAUVALIDO").AsInteger
		  P_EVENTO = CurrentQuery.FieldByName("EVENTO").AsInteger
		  P_LOCALDEEXECUCAO = CurrentQuery.FieldByName("LOCALDEEXECUCAO").AsInteger
		  P_XTHM = CurrentQuery.FieldByName("CODIGOXTHM").AsInteger
		  P_CODPAGAMENTO = CurrentQuery.FieldByName("CODIGOPAGAMENTO").AsInteger
		  P_QUANTIDADE = CurrentQuery.FieldByName("QUANTIDADE").AsInteger
		  P_BENEFICIARIO = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
		  P_LOCALATENDIMENTO = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
		  P_CONDICAOATENDIMENTO = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
		  P_REGIMEATENDIMENTO = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
		  P_TIPOTRATAMENTO = CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger
		  P_OBJETIVOTRATAMENTO = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
		  P_FINALIDADEATENDIMENTO = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
		  P_CBOS = CurrentQuery.FieldByName("CBOS").AsInteger
		  P_RECEBEDOR = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
		  P_TABXTHM = -1
		  P_CONVENIO = CurrentQuery.FieldByName("CONVENIO").AsInteger
		  P_FILIAL = CurrentQuery.FieldByName("FILIAL").AsInteger
		  P_ESTADO = CurrentQuery.FieldByName("ESTADO").AsInteger
		  P_MUNICIPIO = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
		  P_PLANO = CurrentQuery.FieldByName("PLANO").AsInteger
		  P_EXECUTOR = CurrentQuery.FieldByName("EXECUTOR").AsInteger
		  P_ACOMODACAO = CurrentQuery.FieldByName("ACOMODACAO").AsInteger
		  P_TECNICACIRURGICA = CurrentQuery.FieldByName("TECNICAUTILIZADA").AsString

		  If WebMode Then
	        Set Relatorio = CreateBennerObject("Preco.ImprimirPrecoEvento")
			res =	Relatorio.ImprimirPrecoEventoCSharp( P_DATA, P_GRAU, P_EVENTO, _
		                    P_RECEBEDOR, P_LOCALDEEXECUCAO, _
		                    P_XTHM, P_CODPAGAMENTO, P_QUANTIDADE, _
		                    P_BENEFICIARIO, P_LOCALATENDIMENTO, _
		                    P_CONDICAOATENDIMENTO, P_REGIMEATENDIMENTO, _
		                    P_TIPOTRATAMENTO, P_OBJETIVOTRATAMENTO, _
		                    P_FINALIDADEATENDIMENTO, P_CBOS, P_TABXTHM,P_CONVENIO,P_ACOMODACAO,P_ESTADO,P_FILIAL,P_MUNICIPIO,P_PLANO,P_EXECUTOR, _
		                    P_TECNICACIRURGICA)
		    bsShowMessage(res, "I")
		    Set Relatorio = Nothing
		  Else
			  Set Relatorio = CreateBennerObject("Preco.ImprimirPrecoEvento")
		  		If Relatorio.Exec(CurrentSystem, P_DATA, P_GRAU, P_EVENTO, _
		                    CurrentQuery.FieldByName("RECEBEDOR").AsInteger, P_LOCALDEEXECUCAO, _
		                    P_XTHM, P_CODPAGAMENTO, P_QUANTIDADE, _
		                    P_BENEFICIARIO, P_LOCALATENDIMENTO, _
		                    P_CONDICAOATENDIMENTO, P_REGIMEATENDIMENTO, _
		                    P_TIPOTRATAMENTO, P_OBJETIVOTRATAMENTO, _
		                    P_FINALIDADEATENDIMENTO, P_CBOS, P_TECNICACIRURGICA, VsMensagem, vs_PRECOTOTAL, vsMensagemErro) > 0 Then
		    		bsShowMessage(vsMensagemErro, "I")
		  		End If



		  End If
		  Set Relatorio = Nothing
		End Sub

		Public Sub BOTAOLIMPARFILTRO_OnClick()

		  CurrentQuery.FieldByName("GRAUVALIDO").Clear
		  CurrentQuery.FieldByName("EVENTO").Clear
		  CurrentQuery.FieldByName("RECEBEDOR").Clear
		  CurrentQuery.FieldByName("LOCALDEEXECUCAO").Clear
		  CurrentQuery.FieldByName("ESTADO").Clear
		  CurrentQuery.FieldByName("MUNICIPIO").Clear
		  CurrentQuery.FieldByName("BENEFICIARIO").Clear
		  CurrentQuery.FieldByName("PLANO").Clear
		  CurrentQuery.FieldByName("TABHORARIOESPECIAL").AsInteger = 1
		  CurrentQuery.FieldByName("CBOS").Clear

		  If VisibleMode Then
		    VALOREVENTO.Text = ""
		    VALORPF.Text = ""
		  End If

		End Sub

		Public Sub BOTAOVISUALIZAR_OnClick()

		Dim Ordem As Object
		Dim vvContainer As CSDContainer
		Dim vsMensagemErro As String

		Set vvContainer = NewContainer

		  If VisibleMode Then

			Set Ordem = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
		   	Ordem.Exec(CurrentSystem, _
		               1, _
		               "TV_FORM0020", _
		               "Consulta Preço do Evento", _
		               0, _
		               490, _
		               260, _
		               False, _
		               vsMensagemErro, _
		               vvContainer)

		    Set Ordem = Nothing

		  End If
		End Sub


		Public Sub EVENTO_OnPopup(ShowPopup As Boolean)
			Dim procuraDll As Object
			Dim handle As Long

			Set procuraDll = CreateBennerObject("PROCURA.Procurar")

			ShowPopup = False

			handle = procuraDll.ExecTge(CurrentSystem, _
									 "SAM_TGE|*SAM_CBHPM c[c.HANDLE = SAM_TGE.CBHPMTABELA]", _
									 "SAM_TGE.ESTRUTURA|c.ESTRUTURA|SAM_TGE.Z_DESCRICAO|c.DESCRICAO", _
									 1, _
									 "Estrutura TGE|Estrutura CBHPM|Descrição TGE|Descrição CBHPM", _
									 "SAM_TGE.ULTIMONIVEL = 'S'", _
									 "Tabela de Eventos", _
									 False, _
									 EVENTO.LocateText)

			If (handle <> 0) Then
				CurrentQuery.FieldByName("EVENTO").AsInteger = handle
			End If

		    Dim qEvento As Object
			Set qEvento = NewQuery

			qEvento.Add("SELECT GRAUPRINCIPAL FROM SAM_TGE WHERE HANDLE = :PHANDLE")
			qEvento.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
			qEvento.Active = True

			If Not qEvento.FieldByName("GRAUPRINCIPAL").IsNull Then
				CurrentQuery.FieldByName("GRAUVALIDO").Value  = qEvento.FieldByName("GRAUPRINCIPAL").Value
			Else
				bsShowMessage("Não existe grau válido para o evento!","E")
			End If

			Set qEvento = Nothing

			Set procuraDll = Nothing
		End Sub

		Public Sub GRAUVALIDO_OnPopup(ShowPopup As Boolean)

			If CurrentQuery.FieldByName("EVENTO").AsInteger  > 0 Then

				Dim query As Object
				Set query = NewQuery

				query.Add("select count(*) qtde from sam_tge_grau where evento = :phandle")
				query.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
				query.Active = True

				If query.FieldByName("QTDE").AsInteger = 0 Then
		    	    bsShowMessage("Nenhum grau válido para o evento!","E")
					ShowPopup = False
				End If

				Set query = Nothing
			Else
			    bsShowMessage("Evento inválido!","E")
			    ShowPopup = False
		  End If
		End Sub

		Public Sub PLANO_OnPopup(ShowPopup As Boolean)

		  If CurrentQuery.FieldByName("BENEFICIARIO").IsNull  Then
		  	  ShowPopup = False
		     bsShowMessage("Antes de selecionar o plano escolha um beneficiário", "E")
		  Else

			  Dim SQL As Object
			  Set SQL = NewQuery


			  If VisibleMode Then
				  PLANO.LocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_MOD CM JOIN SAM_BENEFICIARIO_MOD BM ON (CM.HANDLE = BM.MODULO) WHERE BM.BENEFICIARIO = " & CurrentQuery.FieldByName("BENEFICIARIO").AsString & ")"
			  Else
				  PLANO.WebLocalWhere = "HANDLE IN (SELECT PLANO FROM SAM_CONTRATO_MOD CM JOIN SAM_BENEFICIARIO_MOD BM ON (CM.HANDLE = BM.MODULO) WHERE BM.BENEFICIARIO = " & CurrentQuery.FieldByName("BENEFICIARIO").AsString & ")"
			  End If

			  Set SQL = Nothing
		  End If

		End Sub



		Public Sub RECEBEDOR_OnPopup(ShowPopup As Boolean)
		  Dim procuraDll As Object
		  Dim viResult As Long
		  Dim VsMensagem As String
		  Dim vlHPrestador As Long

		  Set procuraDll = CreateBennerObject("BSINterface0001.BuscaPrestador")
		  ShowPopup = False

		  viResult = procuraDll.Abrir(CurrentSystem, VsMensagem, 1, RECEBEDOR.LocateText, "T", vlHPrestador)

		  If (viResult > 0) Then
		    If VsMensagem <> "" Then
		      bsShowMessage(VsMensagem, "E")
		    End If
		  Else
		    If vlHPrestador > 0 Then
		      CurrentQuery.FieldByName("RECEBEDOR").AsInteger = vlHPrestador
		    End If
		  End If

		  Set procuraDll = Nothing

		End Sub


		Public Sub TABLE_AfterInsert()

		 Dim QueryGeral As BPesquisa
		 Dim QueryModeloGuia As BPesquisa
		 Dim QueryPais As BPesquisa
		 Dim QueryConvenio As BPesquisa
		 Set QueryModeloGuia = NewQuery
		 Set QueryGeral = NewQuery
		 Set QueryPais  = NewQuery
		 Set QueryConvenio = NewQuery

		 QueryGeral.Clear
		 QueryGeral.Add("SELECT PPC.CODIGOXTHM, PPC.ACOMODACAOAMBULATORIAL, PPC.CODIGOPAGTO")
		 QueryGeral.Add(" FROM SAM_PARAMETROSPROCCONTAS PPC")
		 QueryGeral.Active = True

		 QueryConvenio.Clear
		 QueryConvenio.Add("SELECT HANDLE CONVENIO")
		 QueryConvenio.Add(" FROM SAM_CONVENIO ")
		 QueryConvenio.Add(" WHERE CONVENIOMESTRE = HANDLE")
		 QueryConvenio.Active = True

		 QueryModeloGuia.Clear
		 QueryModeloGuia.Add(" SELECT TMG.CONDICAOATENDIMENTO, TMG.LOCALATENDIMENTO, TMG.REGIMEATENDIMENTO, ")
		 QueryModeloGuia.Add("        TMG.OBJETIVOTRATAMENTO, TMG.TIPOTRATAMENTO, TMG.FINALIDADEATENDIMENTO ")
		 QueryModeloGuia.Add("   FROM SAM_TIPOGUIA_MDGUIA TMG")
		 QueryModeloGuia.Active = True

		 QueryPais.Clear
		 QueryPais.Add("SELECT E.PAIS FROM EMPRESAS E")
		 QueryPais.Active = True

		 CurrentQuery.FieldByName("DATA").AsDateTime = ServerDate
		 CurrentQuery.FieldByName("CODIGOXTHM").AsInteger = QueryGeral.FieldByName("CODIGOXTHM").AsInteger
		 CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger = QueryModeloGuia.FieldByName("CONDICAOATENDIMENTO").AsInteger
		 CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger = QueryModeloGuia.FieldByName("LOCALATENDIMENTO").AsInteger
		 CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger = QueryModeloGuia.FieldByName("REGIMEATENDIMENTO").AsInteger
		 CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger = QueryModeloGuia.FieldByName("OBJETIVOTRATAMENTO").AsInteger
		 CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger = QueryModeloGuia.FieldByName("TIPOTRATAMENTO").AsInteger
		 CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger = QueryModeloGuia.FieldByName("FINALIDADEATENDIMENTO").AsInteger
		 CurrentQuery.FieldByName("ACOMODACAO").AsInteger = QueryGeral.FieldByName("ACOMODACAOAMBULATORIAL").AsInteger
		 CurrentQuery.FieldByName("CODIGOPAGAMENTO").AsInteger = QueryGeral.FieldByName("CODIGOPAGTO").AsInteger
		 CurrentQuery.FieldByName("PAIS").AsInteger = QueryPais.FieldByName("PAIS").AsInteger
		 CurrentQuery.FieldByName("CONVENIO").AsInteger = QueryConvenio.FieldByName("CONVENIO").AsInteger



		 Set QueryModeloGuia = Nothing
		 Set QueryGeral      = Nothing
		 Set QueryPais       = Nothing
		 Set QueryConvenio  = Nothing
		End Sub

		Public Sub TABLE_AfterScroll()


		If (WebMode) Then
		  GRAUVALIDO.WebLocalWhere = " A.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @CAMPO(EVENTO) )"
		Else
		  IMPRIMIR.Visible = False
		  GRAUVALIDO.LocalWhere = " SAM_GRAU.HANDLE IN (SELECT GRAU FROM SAM_TGE_GRAU WHERE EVENTO = @EVENTO )"
		End If

		End Sub

		Public Sub TABLE_BeforePost(CanContinue As Boolean)

		Dim interface As Object
		Dim viRetorno As Long
		Dim P_DATA As Date
		Dim P_HORAATENDIMENTO As Date
		Dim P_GRAU As Long
		Dim P_EVENTO As Long
		Dim P_TABXTHM As Long
		Dim P_ACOMODACAO As Long
		Dim P_XTHM As Long
		Dim P_CODPAGAMENTO As Long
		Dim P_QUANTIDADE As Long
		Dim P_RECEBEDOR As Long
		Dim P_LOCALEXECUCAO As Long
		Dim P_FILIAL As Long
		Dim P_ESTADO As Long
		Dim P_MUNICIPIO As Long
		Dim P_CONVENIO As Long
		Dim P_BENEFICIARIO As Long
		Dim P_PLANO As Long
		Dim P_EXECUTOR As Long
		Dim P_LOCALATENDIMENTO As Long
		Dim P_CONDICAOATENDIMENTO As Long
		Dim P_REGIMEATENDIMENTO As Long
		Dim P_TIPOTRATAMENTO As Long
		Dim P_OBJETIVOTRATAMENTO As Long
		Dim P_FINALIDADEATENDIMENTO As Long
		Dim P_CBOS As Long
		Dim P_IMPRIMIR As Boolean
		Dim P_HORAIOESPECIAL As Boolean
		Dim P_TECNICACIRURGICA As String

		'Atribuir o Valor da CurrentQuery aos Parametros
		If CurrentQuery.FieldByName("TABHORARIOESPECIAL").AsInteger = 1 Then
			P_DATA = CurrentQuery.FieldByName("DATA").AsDateTime
			P_HORAATENDIMENTO = Now
			P_HORAIOESPECIAL = False
		Else
			P_DATA = CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime
			P_HORAATENDIMENTO = CurrentQuery.FieldByName("HORAATENDIMENTO").AsDateTime
			P_HORAIOESPECIAL = True
		End If
		P_GRAU = CurrentQuery.FieldByName("GRAUVALIDO").AsInteger
		P_TECNICACIRURGICA = CurrentQuery.FieldByName("TECNICAUTILIZADA").AsString
		P_EVENTO = CurrentQuery.FieldByName("EVENTO").AsInteger
		P_TABXTHM = -1
		P_ACOMODACAO = CurrentQuery.FieldByName("ACOMODACAO").AsInteger
		P_XTHM = CurrentQuery.FieldByName("CODIGOXTHM").AsInteger
		P_CODPAGAMENTO = CurrentQuery.FieldByName("CODIGOPAGAMENTO").AsInteger
		P_QUANTIDADE = CurrentQuery.FieldByName("QUANTIDADE").AsInteger
		P_RECEBEDOR = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
		P_LOCALEXECUCAO = CurrentQuery.FieldByName("LOCALDEEXECUCAO").AsInteger
		P_FILIAL = CurrentQuery.FieldByName("FILIAL").AsInteger
		P_ESTADO = CurrentQuery.FieldByName("ESTADO").AsInteger
		P_MUNICIPIO = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
		P_CONVENIO = CurrentQuery.FieldByName("CONVENIO").AsInteger
		P_BENEFICIARIO = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
		P_PLANO = CurrentQuery.FieldByName("PLANO").AsInteger
		P_EXECUTOR = CurrentQuery.FieldByName("EXECUTOR").AsInteger
		P_LOCALATENDIMENTO = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
		P_CONDICAOATENDIMENTO = CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsInteger
		P_REGIMEATENDIMENTO = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
		P_TIPOTRATAMENTO = CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger
		P_OBJETIVOTRATAMENTO = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
		P_FINALIDADEATENDIMENTO = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
		P_CBOS = CurrentQuery.FieldByName("CBOS").AsInteger
		P_IMPRIMIR = CurrentQuery.FieldByName("IMPRIMIR").AsBoolean

		If P_CONVENIO <= 0 And P_BENEFICIARIO <= 0 Then
		  bsShowMessage("É necessário preencher o Convênio ou o Beneficiário","E")
		  CanContinue = False
		  Exit Sub
		End If

		If P_EVENTO <= 0 Or P_GRAU <= 0 Then
		  bsShowMessage("Evento e Grau não podem ficar nulos","E")
		  CanContinue = False
		  Exit Sub
		End If

		  Set interface =CreateBennerObject("PRECO.CONSULTAPRECOEVENTO")
		  viRetorno = interface.Exec(CurrentSystem, P_DATA, P_GRAU, P_EVENTO, P_TABXTHM, _
		                     P_ACOMODACAO, P_XTHM, P_CODPAGAMENTO, P_QUANTIDADE, _
		                     P_RECEBEDOR, P_LOCALEXECUCAO, P_FILIAL, P_ESTADO, P_MUNICIPIO, _
		                     P_CONVENIO, P_BENEFICIARIO, P_PLANO, P_EXECUTOR, _
		                     P_LOCALATENDIMENTO, P_CONDICAOATENDIMENTO, _
		                     P_REGIMEATENDIMENTO, P_TIPOTRATAMENTO, _
		                     P_OBJETIVOTRATAMENTO, P_FINALIDADEATENDIMENTO, P_CBOS, P_HORAATENDIMENTO, P_HORAIOESPECIAL, P_TECNICACIRURGICA, _
		                     vs_PRECOTOTAL, vs_VALORPF, VsMensagem)

		  Set interface = Nothing

		If WebMode Then
			If P_IMPRIMIR Then
				BOTAOIMPRIMIR_OnClick
		  	Exit Sub
		  End If
		End If


		  If (viRetorno > 0) Then
		    bsShowMessage(VsMensagem, "E")
		    CanContinue = False
		  Else
		    If VisibleMode Then
		      VALOREVENTO.Text = vs_PRECOTOTAL
		      VALORPF.Text = vs_VALORPF
		      CanContinue = False
		    Else
		      bsShowMessage("VALOR TOTAL = " + vs_PRECOTOTAL +  Chr(13)  + "VALOR DA PF = " + vs_VALORPF,"I")
		    End If
		  End If

		End Sub

		Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
			If CommandID = "CMD_IMPRIMIR_RELATORIO" Then
				BOTAOIMPRIMIR_OnClick
			End If
		End Sub


		Public Sub TABLE_NewRecord()
		   'Artur - SMS 93336 - 07-03-2008
		   If (SessionVar("HANDLECA005")<>"") Then
		      CurrentQuery.FieldByName("RECEBEDOR").AsInteger = CLng(SessionVar("HANDLECA005"))
		   End If
		   'Crislei - SMS 108068 - 10-02-2008
		   If (SessionVar("HANDLECA006")<>"") Then
		   	  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = CLng(SessionVar("HANDLECA006"))
		   'Gustavo Galina - SMS 131975 - 04-06-2010
		   ElseIf (SessionVar("HANDLESAMINCOMP")<>"") Then
		   	  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = CLng(SessionVar("HANDLESAMINCOMP"))
		   End If
		End Sub
