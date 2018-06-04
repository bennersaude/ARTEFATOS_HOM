'HASH: 7B7825D17BA4A128AC72181FBCAF91B7
 
 Dim qConsulta   As BPesquisa
 Dim q 		   	 As BPesquisa
 Dim dtImpressao As BPesquisa
 Dim impressao 	 As Integer

Public Sub imprimir

	impressao = 0

			Dim qt As Object
			Set qt = NewQuery

			qt.Add("SELECT HANDLE FROM R_RELATORIOS WHERE UPPER(CODIGO) = 'GTO001' ")
			qt.Active = True


				If Not qt.EOF Then

			  		Dim rep1 As CSReportPrinter
			  		Set rep1 = NewReport(qt.FieldByName("HANDLE").AsInteger)

			  		SessionVar("ProtocoloCapaPeg") = CurrentQuery.FieldByName("PEG").AsString

			  		rep1.Preview
			  		Set rep1 = Nothing

			  		InfoDescription = "Formulário de capa de lote de PEG impresso com sucesso!"
			  		impressao = 1

					If impressao = 1 Then

					         Dim dDataAgora As Date
					 		  dDataAgora = ServerNow

					    If (Not InTransaction) Then
					     	StartTransaction
						End If

						  Set dtImpressao  = NewQuery
						  	  dtImpressao.Clear
							  dtImpressao.Add("UPDATE GTO_PROT_COBRANCA SET DATAIMPRESSAO = :DTIMPRESSAO WHERE PEG = :PEG")
							  dtImpressao.ParamByName("PEG").AsInteger = CurrentQuery.FieldByName("PEG").AsInteger
							  dtImpressao.ParamByName("DTIMPRESSAO").AsDateTime = dDataAgora
							  dtImpressao.ExecSQL

						If InTransaction Then
						  Commit
						End If

					End If



			  	Else
			  		InfoDescription = "Relatório não encontrado."
			  		CanContinue = False
			  		impressao = 0

				End If

			Set qt = Nothing
			Set dtImpressao = Nothing

End Sub

Public Sub TABLE_AfterScroll()


	 		Set qConsulta = NewQuery

				qConsulta.Active = False
				qConsulta.Clear
				qConsulta.Add("SELECT COUNT (DISTINCT GUIA.HANDLE) AS QT_GUIA													  ")
				qConsulta.Add(", SUM (EV.VALORAPRESENTADO) As VALORINFORMADO													  ")
				qConsulta.Add(", PEG.PEG                                                                                          ")
				qConsulta.Add("		FROM GTO_PROT_COBRANCA As PCOB INNER JOIN SAM_PEG As PEG On PCOB.PEG = PEG.PEG         ")
				qConsulta.Add("			          						  INNER Join SAM_GUIA As GUIA On GUIA.PEG = PEG.Handle    ")
				qConsulta.Add("		                                      INNER Join SAM_GUIA_EVENTOS EV  On EV.GUIA = GUIA.Handle")
				qConsulta.Add("		WHERE PCOB.HANDLE = :HPEG")
				qConsulta.Add("		GROUP BY PEG.PEG																			  ")
				qConsulta.ParamByName("HPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				qConsulta.Active = True

			Set qAlteracao = NewQuery

				qAlteracao.Active = False
				qAlteracao.Clear
				qAlteracao.Text = "SELECT QTDGUIAINFORMADA,TOTALINFORMADO FROM GTO_PROT_COBRANCA WHERE HANDLE = :HPEG"
				qAlteracao.ParamByName("HPEG").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
				qAlteracao.RequestLive = True
				qAlteracao.Active = True

			    ' Altera o registro na base de dados
				qAlteracao.Edit
				qAlteracao.FieldByName("QTDGUIAINFORMADA").AsInteger = qConsulta.FieldByName("QT_GUIA").AsFloat
				qAlteracao.FieldByName("TOTALINFORMADO").AsFloat = qConsulta.FieldByName("VALORINFORMADO").AsFloat

				qAlteracao.Post


			Set qConsulta = Nothing
			Set qAlteracao = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

	If CommandID = "IMPRIMIR" Then

			Dim qPesquisa As Object
			Set qPesquisa = NewQuery
				qPesquisa.Clear
				qPesquisa.Add("SELECT REDEDIFERENCIADA FROM SAM_PRESTADOR WHERE HANDLE = :CODPRESTADOR")
				qPesquisa.ParamByName("CODPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
				qPesquisa.Active = True

			If qPesquisa.FieldByName("REDEDIFERENCIADA").AsString = "S" And (CurrentQuery.FieldByName("NF").AsInteger = 0) Or (CurrentQuery.FieldByName("DATAEMISSAONF").AsDateTime = Null) Then

				CancelDescription = "É Obrigatório informar os dados da nota fiscal"
				CanContinue = False

			Else

				If CurrentQuery.FieldByName("SITUACAO").AsString <> "E" Then

 				CancelDescription = "A impressão do relatório não está disponível."

 				CanContinue = False

				Else

		 	   		Call imprimir
		 	   	End If

	 		End If

	 		Set qPesquisa = Nothing

	End If

End Sub
