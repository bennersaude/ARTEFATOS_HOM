'HASH: 5D0856713BB428C5D11CDD5526E970C8
'Macro: SFN_DOCUMENTO
'#Uses "*bsShowMessage
Option Explicit

Public Sub BOTAOALTERARVENCIMENTO_OnClick()
	If Not CurrentQuery.FieldByName("BAIXADATA").IsNull Then
		bsShowMessage("Documento já baixado", "I")

		Exit Sub
	End If

	If (PermissionFieldByHandle(HandleOfField(HandleOfTable("SFN_DOCUMENTO"), "DATAVENCIMENTO")) <= 1) Then
    	bsShowMessage("Usuário sem premissão para alterar a data de vencimento do documento.", "I")

		Exit Sub
	End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
	Dim SQL As BPesquisa
	Dim INTERFACE0002 As Object
	Dim vsMensagem As String
	Dim vcContainer As CSDContainer
	Set SQL = NewQuery
	Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

	SQL.Clear
	SQL.Active = False
	SQL.Add("SELECT D.NUMERO ")
	SQL.Add("  FROM SFN_DOCUMENTO D, ")
	SQL.Add("		SFN_ROTINAARQUIVO_DOC RAD ")
	SQL.Add(" WHERE D.ULTIMAROTINAARQUIVODOC = RAD.HANDLE ")
	SQL.Add("	AND RAD.TABENVIORETORNO		 = 1 ")
	SQL.Add("	AND D.HANDLE				 = :DOCUMENTO ")
	SQL.ParamByName("DOCUMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.Active = True

	If Not SQL.EOF Then
		If (bsShowMessage("Documento com rotina arquivo de ENVIO" + Chr(13) + "Deseja continuar ?", "Q") = vbYes) Then
			SessionVar("HDOCUMENTO") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)

			If VisibleMode Then INTERFACE0002.Exec(CurrentSystem, _
												   1, _
												   "TV_FORM0042", _
												   "Cancelamento de documento", _
												   0, _
												   178, _
												   481, _
												   False, _
												   vsMensagem, _
												   vcContainer)
		End If
	Else
		SessionVar("HDOCUMENTO") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)

		If VisibleMode Then INTERFACE0002.Exec(CurrentSystem, _
											   1, _
											   "TV_FORM0042", _
											   "Cancelamento de documento", _
											   0, _
											   178, _
											   481, _
											   False, _
											   vsMensagem, _
											   vcContainer)
	End If

	If VisibleMode Then RefreshNodesWithTable("SFN_DOCUMENTO")

	Set SQL = Nothing
End Sub

Public Sub BOTAOFATURAS_OnClick()
  Dim OLEBaixa As Object
  Set OLEBaixa = CreateBennerObject("SfnBaixa.Documento")

  OLEBaixa.bXdOC(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set OLEBaixa = Nothing
End Sub

Public Sub BOTAOBAIXAR_OnClick()
	Dim SQL As BPesquisa
	Dim Interface As Object
	Dim vsMensagem As String
	Dim vcContainer As CSDContainer

	If Not CurrentQuery.FieldByName("BAIXADATA").IsNull Then
		bsShowMessage("Documento já baixado", "I")

		Exit Sub
	End If

	Set SQL = NewQuery

	SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT SUM(CASE WHEN F.NATUREZA = 'C' THEN -F.SALDO ELSE F.SALDO END) SALDOFATURA ")
    SQL.Add("  FROM SFN_FATURA F ")
    SQL.Add("  JOIN SFN_DOCUMENTO_FATURA DF ON DF.FATURA = F.HANDLE ")
    SQL.Add(" WHERE DF.DOCUMENTO = :DOCUMENTO ")
    SQL.ParamByName("DOCUMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If Abs(Round(SQL.FieldByName("SALDOFATURA").AsFloat,2)) <> Round(CurrentQuery.FieldByName("VALOR").AsFloat, 2) Then
      If bsShowMessage("Documento com valor diferente do saldo das faturas!" + Chr(13) + "Deseja continuar?", "Q") = vbNo Then
		Set SQL = Nothing
        Exit Sub
      End If
    End If


	SQL.Active = False
	SQL.Clear
	SQL.Add("SELECT D.NUMERO")
	SQL.Add("  FROM SFN_DOCUMENTO D,")
	SQL.Add("		SFN_ROTINAARQUIVO_DOC RAD")
	SQL.Add(" WHERE D.ULTIMAROTINAARQUIVODOC = RAD.HANDLE")
	SQL.Add("	AND RAD.TABENVIORETORNO = 1")
	SQL.Add("	AND D.HANDLE = :DOCUMENTO")
	SQL.ParamByName("DOCUMENTO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	SQL.Active = True

	SessionVar("HDOCUMENTO") = CurrentQuery.FieldByName("HANDLE").AsString

	If Not SQL.EOF Then
		If bsShowMessage("Documento com rotina arquivo de ENVIO" + Chr(13) + "Deseja continuar ?", "Q") = vbYes Then
			'If VisibleMode Then
				Set Interface = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

				Interface.Exec(CurrentSystem, _
							   1, _
							   "TV_FORM0043", _
							   "Baixa de documento", _
							   0, _
							   560, _
							   530, _
							   False, _
							   vsMensagem, _
							   vcContainer)

				Set Interface = Nothing
			'End If
		End If
	Else
		'If VisibleMode Then
			Set Interface = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

			Interface.Exec(CurrentSystem, _
						   1, _
						   "TV_FORM0043", _
						   "Baixa de documento", _
						   0, _
						   560, _
						   530, _
						   False, _
						   vsMensagem, _
						   vcContainer)

			Set Interface = Nothing
		'End If
	End If

	Set SQL = Nothing
End Sub

Public Sub BOTAOCONTAFINANCEIRA_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("SamContaFinanceira.Consulta")

  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger)

  Set Interface = Nothing
End Sub

Public Sub BOTAOESTORNOBAIXA_OnClick()
  If CurrentQuery.FieldByName("BAIXADATA").IsNull Then
	bsShowMessage("O documento não está baixado", "I")
  Else
  	SessionVar("HDOCUMENTO") = CurrentQuery.FieldByName("HANDLE").AsString

	If VisibleMode Then
		Dim INTERFACE0002 As Object
		Dim vsMensagem As String
		Dim vcContainer As CSDContainer
		Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

		INTERFACE0002.Exec(CurrentSystem, _
						   1, _
						   "TV_FORM0050", _
						   "Estorno de baixa", _
						   0, _
						   400, _
						   530, _
						   False, _
						   vsMensagem, _
						   vcContainer)

		CurrentQuery.Active = False
		CurrentQuery.Active = True

		Set INTERFACE0002 = Nothing
	End If
  End If
End Sub

Public Sub BOTAOEXCLUIR_OnClick()
  bsShowMessage("Operação não suportada", "I")
  Exit Sub
End Sub

Public Sub BOTAOIMPRIMIR_OnClick()
  Dim vContaFinDono As Long
  Dim vAchou As Boolean
  Dim vLogradouro As Variant
  Dim vNUMERO As Variant
  Dim vCOMPLEMENTO As Variant
  Dim vBAIRRO As Variant
  Dim vCidade As Variant
  Dim vCidadeHandle As Variant
  Dim vESTADO As Variant
  Dim vSigla As Variant
  Dim vCEP As Variant
  Dim vTelResidencial As Variant
  Dim vTelComercial As Variant
  Dim vCelular As Variant
  Dim vFax As Variant
  Dim EnderecoDll As Object
  Dim SQL As Object
  Dim qUpDoc As Object

  If CurrentQuery.State <>1 Then
	bsShowMessage("O Documento não pode estar em edição", "I")
	Exit Sub
  End If

  If Not CurrentQuery.FieldByName("CANCDATA").IsNull Then
	bsShowMessage("Não pode imprimir um documento cancelado", "I")
	Exit Sub
  End If

  Set SQL = NewQuery

  SQL.Clear

  SQL.Add("SELECT TABRESPONSAVEL, BENEFICIARIO, PESSOA FROM SFN_CONTAFIN WHERE HANDLE=:CONTAFIN")

  SQL.ParamByName("CONTAFIN").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABRESPONSAVEL").AsInteger <>2 Then 'Diferente de Prestador,verificar endereço
	If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then 'Beneficiário
	  vContaFinDono = SQL.FieldByName("BENEFICIARIO").AsInteger

	  Set EnderecoDll = CreateBennerObject("SAMENDERECO.LOCALIZA")

	  EnderecoDll.Executar(CurrentSystem, vContaFinDono, 0, "B", "I", vAchou, vLogradouro, vNUMERO, vCOMPLEMENTO, vBAIRRO, _
						   vCidade, vCidadeHandle, vESTADO, vSigla, vCEP, vTelResidencial, vTelComercial, _
						   vCelular, vFax)
	  Set EnderecoDll = Nothing
	Else 'Pessoa
	  vContaFinDono = SQL.FieldByName("PESSOA").AsInteger

	  SQL.Active = False

	  SQL.Clear

	  ' Coelho SMS: 73167 utilizando a mesma lógica que o botão imprimir da ficha financeira
	  SQL.Add("SELECT M.NOME MUNICIPIONOME, E.SIGLA,                                  ")
	  SQL.Add("       EP.BAIRRO EPBAIRRO, EP.CEP EPCEP, EP.COMPLEMENTO EPCOMPLEMENTO, ")
	  SQL.Add("       EP.LOGRADOURO EPLOGRADOURO, EP.NUMERO EPNUMERO                  ")
	  SQL.Add("  FROM SFN_PESSOA C                                                    ")
	  SQL.Add("  JOIN SAM_ENDERECO EP ON (EP.HANDLE = C.ENDERECOCPFCNPJ)              ")
	  SQL.Add("  LEFT JOIN ESTADOS E ON (E.HANDLE= EP.ESTADO)                         ")
	  SQL.Add("  LEFT JOIN MUNICIPIOS M ON (M.HANDLE=EP.MUNICIPIO)                    ")
	  SQL.Add("  WHERE C.HANDLE=:HANDLE                                               ")

	  SQL.ParamByName("HANDLE").AsInteger = vContaFinDono
	  SQL.Active = True

	  vBAIRRO = SQL.FieldByName("EPBAIRRO").AsString
	  vCEP = SQL.FieldByName("EPCEP").AsString
	  vCOMPLEMENTO = SQL.FieldByName("EPCOMPLEMENTO").AsString
	  vLogradouro = SQL.FieldByName("EPLOGRADOURO").AsString
	  vNUMERO = SQL.FieldByName("EPNUMERO").AsInteger
	  vCidade = SQL.FieldByName("MUNICIPIONOME").AsString
	  vSigla = SQL.FieldByName("SIGLA").AsString
	End If

	If (CurrentQuery.FieldByName("BAIRRO").AsString <> vBAIRRO) Or _
	   (CurrentQuery.FieldByName("CEP").AsString <> vCEP) Or _
	   (CurrentQuery.FieldByName("COMPLEMENTO").AsString <> vCOMPLEMENTO) Or _
	   (CurrentQuery.FieldByName("ENDERECO").AsString <> vLogradouro) Or _
	   (CurrentQuery.FieldByName("NUMEROENDERECO").AsInteger <> vNUMERO) Then

      Dim vbAtualizarEndereco As Boolean

      vbAtualizarEndereco = False

      If WebMode Then
        vbAtualizarEndereco = True
        bsShowMessage("Dados de endereço do documento estavam diferentes do endereço atual e foram atualizados!", "I")
      Else
        If bsShowMessage("Dados de endereço foram alterados na conta financeira.  Atualizar ?", "Q") = vbYes Then
          vbAtualizarEndereco = True
        End If
      End If

	  If vbAtualizarEndereco Then
		Set qUpDoc = NewQuery

		qUpDoc.Clear

		If Not InTransaction Then StartTransaction
			qUpDoc.Add("UPDATE SFN_DOCUMENTO")
			qUpDoc.Add("  SET BAIRRO         = :BAIRRO,")
			qUpDoc.Add("      CEP            = :CEP,")
			qUpDoc.Add("      COMPLEMENTO    = :COMPLEMENTO,")
			qUpDoc.Add("      ENDERECO       = :ENDERECO,")
			qUpDoc.Add("      ESTADO         = :ESTADO,")
			qUpDoc.Add("      MUNICIPIO      = :MUNICIPIO,")
			qUpDoc.Add("      NUMEROENDERECO = :NUMEROENDERECO")
			qUpDoc.Add("WHERE HANDLE=:HANDLE")

			qUpDoc.ParamByName("BAIRRO").AsString = Str(vBAIRRO)
			qUpDoc.ParamByName("CEP").AsString = Str(vCEP)
			qUpDoc.ParamByName("COMPLEMENTO").AsString = Str(vCOMPLEMENTO)
			qUpDoc.ParamByName("ENDERECO").AsString = Str(vLogradouro)
			qUpDoc.ParamByName("ESTADO").AsString = Str(vSigla)
			qUpDoc.ParamByName("MUNICIPIO").AsString = Str(vCidade)

			If (vNUMERO = 0) Or (vNUMERO = Null) Then
		  	qUpDoc.ParamByName("NUMEROENDERECO").DataType = ftInteger
		  	qUpDoc.ParamByName("NUMEROENDERECO").Clear
			Else
		  	qUpDoc.ParamByName("NUMEROENDERECO").AsInteger = Int(vNUMERO)
			End If

			qUpDoc.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			qUpDoc.ExecSQL
		If InTransaction Then Commit

		Set qUpDoc = Nothing
  	  Else
		Set SQL = Nothing

		Exit Sub
	  End If
	End If
  End If

  If VisibleMode Then
    Dim Obj As Object
    Set Obj = CreateBennerObject("SamImpressao.Boleto")

    Obj.ImprimirBoleto(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

    Set Obj = Nothing
  Else
    SQL.Clear
    SQL.Add("SELECT RELATORIOIMPRESSAOWEB")
    SQL.Add("FROM SFN_TIPODOCUMENTO")
    SQL.Add("WHERE HANDLE = :HTIPODOCUMENTO")
    SQL.ParamByName("HTIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
    SQL.Active = True

    If SQL.FieldByName("RELATORIOIMPRESSAOWEB").IsNull Then
      bsShowMessage("Relatório de impressão não está configurado no Tipo de Documento!", "I")
    Else
      SessionVar("HDOCUMENTO_IMPRESSAOBOLETO") = CurrentQuery.FieldByName("HANDLE").AsString

      ReportPreview(SQL.FieldByName("RELATORIOIMPRESSAOWEB").AsInteger, "", False, False)
    End If
  End If

  Set SQL = Nothing
End Sub

Public Sub BOTAONOTA_OnClick()
  Dim Interface As Object
  Dim X As Long
  Dim vErro As String
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT D.* FROM SFN_DOCUMENTO D, SFN_NOTA_DOCUMENTO ND")
  SQL.Add("WHERE D.HANDLE = ND.DOCUMENTO")
  SQL.Add("AND D.HANDLE= " + CurrentQuery.FieldByName("HANDLE").AsString)

  SQL.Active = True

  If(CurrentQuery.FieldByName("CANCDATA").AsString = "")Then
	If SQL.EOF Then
	  If CurrentQuery.FieldByName("NATUREZA").AsString = "C" Then
		Set Interface = CreateBennerObject("SFNNOTA.ROTINAS")

		Interface.GeraNotaDoc(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, "S", vErro)

		Set Interface = Nothing
	  Else
		bsShowMessage("Geração de nota cancelada, documento com natureza de débito", "I")
	  End If
	Else
	  bsShowMessage("Já foi Gerada Nota Fiscal para este Documento", "I")
	End If
  Else
	bsShowMessage("Documento Baixado ou Cancelado! ", "I")
  End If
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SessionVar("WebHandleDocumento") = CurrentQuery.FieldByName("HANDLE").AsString 'Luciano T. Alberti - SMS 85992 - 22/10/2007

  SQL.Clear

  SQL.Active = False

  SQL.Add("SELECT INSTRUCAO1, INSTRUCAO2, INSTRUCAO3, INSTRUCAO4, INSTRUCAO5")
  SQL.Add("FROM SFN_TIPODOCUMENTO WHERE HANDLE = :HTIPODOC")

  SQL.ParamByName("HTIPODOC").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("INSTRUCAO1").AsString <>"" Then
	INSTRUCAO1.ReadOnly = True
  Else
	INSTRUCAO1.ReadOnly = False
  End If

  If SQL.FieldByName("INSTRUCAO2").AsString <>"" Then
	INSTRUCAO2.ReadOnly = True
  Else
	INSTRUCAO2.ReadOnly = False
  End If

  If SQL.FieldByName("INSTRUCAO3").AsString <>"" Then
	INSTRUCAO3.ReadOnly = True
  Else
	INSTRUCAO3.ReadOnly = False
  End If

  If SQL.FieldByName("INSTRUCAO4").AsString <>"" Then
	INSTRUCAO4.ReadOnly = True
  Else
	INSTRUCAO4.ReadOnly = False
  End If

  If SQL.FieldByName("INSTRUCAO5").AsString <>"" Then
	INSTRUCAO5.ReadOnly = True
  Else
	INSTRUCAO5.ReadOnly = False
  End If

  Set SQL = Nothing

  BOTAOCANCELAR.Enabled = (CurrentQuery.FieldByName("BAIXADATA").IsNull And CurrentQuery.FieldByName("CANCDATA").IsNull)
  BOTAOBAIXAR.Enabled = (CurrentQuery.FieldByName("BAIXADATA").IsNull And CurrentQuery.FieldByName("CANCDATA").IsNull)
  BOTAOESTORNOBAIXA.Enabled = (Not CurrentQuery.FieldByName("BAIXADATA").IsNull)
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOCANCELAR"
	  BOTAOCANCELAR_OnClick
	Case "BOTAOFATURAS"
	  BOTAOFATURAS_OnClick
	Case "BOTAOBAIXAR"
	  BOTAOBAIXAR_OnClick
	Case "BOTAOCONTAFINANCEIRA"
	  BOTAOCONTAFINANCEIRA_OnClick
	Case "BOTAOESTORNOBAIXA"
	  BOTAOESTORNOBAIXA_OnClick
	Case "BOTAOEXCLUIR"
	  BOTAOEXCLUIR_OnClick
	Case "BOTAOIMPRIMIR"
	  BOTAOIMPRIMIR_OnClick
	Case "BOTAONOTA"
	  BOTAONOTA_OnClick
	Case "BOTAOALTERARVENCIMENTO"
	  BOTAOALTERARVENCIMENTO_OnClick
  End Select
End Sub
