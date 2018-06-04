'HASH: 577C920595E38B7E6155742F66B241DC
'Macro: SFN_FATURA
'#Uses "*bsShowMessage"

Option Explicit


Private Sub EfetivaCancelamento()
	SessionVar("HFATURA") = CurrentQuery.FieldByName("HANDLE").AsString

	If (bsShowMessage("Confirma cancelamento?", "Q") = vbYes) Then
		If VisibleMode Then
			Dim INTERFACE0002 As Object
			Dim vsMensagem As String
	        Dim SAMCONTAFINANCEIRA As Object
	        Dim vsDocumentos As String
	        Dim vsMensagemRetorno As String

			Dim vcContainer As CSDContainer

			Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
			Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")

			If SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,vsDocumentos,vsMensagemRetorno) Then

			  INTERFACE0002.Exec(CurrentSystem, _
			  				     1, _
			 				     "TV_FORM0041", _
							     "Cancelamento de fatura", _
							     0, _
							     482, _
							     481, _
							     False, _
							     vsMensagem, _
							     vcContainer)

	          Set INTERFACE0002 = Nothing
	        Else
	          BsShowMessage(vsMensagemRetorno,"E")
	        End If

	        Set SAMCONTAFINANCEIRA = Nothing

		End If
	End If

	If VisibleMode Then RefreshNodesWithTable("SFN_FATURA")
End Sub

Public Sub BOTAOATUALIZAR_OnClick()
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
	Exit Sub
  End If

  Set Sql = Nothing
  Dim Interface As Object
  Set Interface = CreateBennerObject("FINANCEIRO.Fatura")
  Interface.Atualiza(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing
  CurrentQuery.Active = False
  CurrentQuery.Active = True
  RefreshNodesWithTable("SFN_FATURA")
End Sub

Public Sub BOTAOBAIXAR_OnClick()
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
	Exit Sub
  End If

  Set Sql = Nothing

  SessionVar("HFATURA") = CurrentQuery.FieldByName("HANDLE").AsString

  Dim INTERFACE0002 As Object
  Dim vsMensagem As String

  Dim SAMCONTAFINANCEIRA As Object
  Dim vsDocumentos As String
  Dim vsMensagemRetorno As String

  Dim vcContainer As CSDContainer

  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")

  If SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,vsDocumentos,vsMensagemRetorno) Then

	INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0038", _
					   "Baixa de fatura",  _
					   0, _
					   560, _
					   530, _
					   False, _
					   vsMensagem, _
					   vcContainer)

	Set INTERFACE0002 = Nothing
  Else
    BsShowMessage(vsMensagemRetorno,"I")
  End If

  Set SAMCONTAFINANCEIRA = Nothing

  If VisibleMode Then
	RefreshNodesWithTable("SFN_FATURA")
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
	Exit Sub
  End If

  Sql.Active = False
  Sql.Clear
  Sql.Add("SELECT COUNT(1) LANCANTEC                           ")
  Sql.Add("  FROM SFN_FATURA_LANC L                            ")
  Sql.Add("  JOIN SIS_OPERACAO    O ON (O.HANDLE = L.OPERACAO) ")
  Sql.Add(" WHERE O.CODIGO IN ('111', '112')                   ")
  Sql.Add("   AND L.FATURA = :HFATURA                          ")

  Sql.ParamByName("HFATURA").AsString = CurrentQuery.FieldByName("HANDLE").AsString
  Sql.Active = True

  If Sql.FieldByName("LANCANTEC").AsInteger > 0 Then
    Set Sql = Nothing
	If (bsShowMessage("Processo é irreversível para fatura com lançamento de pagamento antecipado. Confirma cancelamento?", "Q") = vbYes) Then
	  EfetivaCancelamento
	End If
  Else
    Set Sql = Nothing
	EfetivaCancelamento
  End If
End Sub

Public Sub BOTAOCONTAFINANCEIRA_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("SamContaFinanceira.Consulta")
  Interface.Exec(CurrentSystem, CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger)
  Set Interface = Nothing

End Sub

Public Sub BOTAOEXCLUIR_OnClick()
  If bsShowMessage("Confirma Exclusão?", "Q") = vbYes Then
    Dim Interface As Object
    Set Interface = CreateBennerObject("FINANCEIRO.Fatura")
    If Interface.Excluir(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)<= 0 Then
      bsShowMessage("Erro na exclusão!", "I")
    Else
      bsShowMessage("Excluído com sucesso!", "I")
      CurrentQuery.Active = False
      CurrentQuery.Active = True
      RefreshNodesWithTable("SFN_FATURA")
    End If
    Set Interface = Nothing
  End If
End Sub


Public Sub BOTAOCOPIAR_OnClick()
  Dim Sql1 As Object

  Set Sql1 = NewQuery
  Sql1.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
  Sql1.Active = True

  If Sql1.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
	Exit Sub
  End If

  Set Sql1 = Nothing

  Dim Interface As Object
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT HANDLE FROM SFN_DOCUMENTO_FATURA WHERE FATURA = " + CurrentQuery.FieldByName("HANDLE").AsString)
  Sql.Active = True
  If Sql.EOF Then
    Set Interface = CreateBennerObject("SFNFatura.Rotinas")
    Interface.Copiar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing
  Else
    bsShowMessage("Existe documento aberto para esta fatura. Não é possível continuar", "I")
    Exit Sub
  End If
End Sub

Public Sub BOTAOESTORNOBAIXA_OnClick()
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
	Exit Sub
  End If

  Set Sql = Nothing

	SessionVar("HFATURA") = CurrentQuery.FieldByName("HANDLE").AsString

	Dim qRotArq As BPesquisa
	Dim qParamFin As BPesquisa
	Set qRotArq = NewQuery
	Set qParamFin = NewQuery

	qRotArq.Clear
	qRotArq.Add("SELECT L.TESOURARIALANC, L.ROTINAARQUIVO")
	qRotArq.Add("  FROM SFN_FATURA_LANC L, SIS_OPERACAO O")
	qRotArq.Add(" WHERE O.HANDLE = L.OPERACAO")
	qRotArq.Add("	AND L.FATURA = :HFATURA")
	qRotArq.Add("	AND O.CODIGO = :CODIGO")
	qRotArq.ParamByName("HFATURA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qRotArq.ParamByName("CODIGO").AsInteger = 130
	qRotArq.Active = True

	qParamFin.Clear
	qParamFin.Add("SELECT PERMITEESTORNOBAIXAROTARQ FROM SFN_PARAMETROSFIN")
	qParamFin.Active = True

	If (qRotArq.FieldByName("ROTINAARQUIVO").AsInteger > 0) And (Not (qParamFin.FieldByName("PERMITEESTORNOBAIXAROTARQ").AsString = "S")) Then
		bsShowMessage("Lançamento vinculado a uma rotina arquivo, não pode ser estornado!", "I")

		Exit Sub
	End If

	If (qRotArq.FieldByName("TESOURARIALANC").IsNull Or (qRotArq.FieldByName("TESOURARIALANC").AsInteger = 0)) Then
		If VisibleMode Then
			Dim INTERFACE0002 As Object
			Dim vsMensagem As String
			Dim vcContainer As CSDContainer

			Dim SAMCONTAFINANCEIRA As Object
            Dim vsDocumentos As String
            Dim vsMensagemRetorno As String

            Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")
            Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

			If SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,vsDocumentos,vsMensagemRetorno) Then

  			  INTERFACE0002.Exec(CurrentSystem, _
			 				     1, _
							     "TV_FORM0051", _
							     "Estorno de baixa", _
							     0, _
							     400, _
							     530, _
							     False, _
							     vsMensagem, _
							     vcContainer)

   			  Set INTERFACE0002 = Nothing
            Else
              BsShowMessage(vsMensagemRetorno,"E")
            End If

            Set SAMCONTAFINANCEIRA = Nothing

		End If
	Else
		bsShowMessage("Fatura com lançamento p/ Tesouraria somente pode ser estornado na Tesouraria!", "I")
	End If
End Sub

Public Sub BOTAOESTORNOCANCELA_OnClick()
  Dim Sql As Object
  Dim vsDocumentos As String
  Dim AUX As Boolean

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
	Exit Sub
  End If

  Set Sql = Nothing

	Dim EspecificoDll As Object
	Set EspecificoDll = CreateBennerObject("ESPECIFICO.uESPECIFICO")

	vsDocumentos = EspecificoDll.FIN_VerificaDocumentoFatura(CurrentSystem, _
															 CurrentQuery.FieldByName("HANDLE").AsInteger)

	If (Not vsDocumentos = "") Then
		If (bsShowMessage("Existe documento aberto para esta fatura. Continuar?", "Q") = vbYes) Then
			AUX = True
		Else
			Exit Sub
		End If
	End If

	SessionVar("HFATURA") = CurrentQuery.FieldByName("HANDLE").AsString

	If VisibleMode Then
		Dim INTERFACE0002 As Object
		Dim vsMensagem As String
		Dim vcContainer As CSDContainer
		Dim SAMCONTAFINANCEIRA As Object
        Dim vsDocs As String
        Dim vsMensagemRetorno As String

		Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
		Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")

        If SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger,vsDocs,vsMensagemRetorno) Then

		  INTERFACE0002.Exec(CurrentSystem, _
						     1, _
						     "TV_FORM0044", _
						     "Estorno de cancelamento", _
						     0, _
						     482, _
						     481, _
						     False, _
						     vsMensagem, _
						     vcContainer)
		  Set INTERFACE0002 = Nothing
		Else
		  BsShowMessage(vsMensagemRetorno,"E")
        End If

        Set SAMCONTAFINANCEIRA = Nothing

	End If
End Sub
Public Sub BOTAOALTERARVENCIMENTO_OnClick()

    If Not CurrentSystem.WebMode Then
		If Not (CurrentSystem.PermissionFieldByName(CurrentSystem.HandleOfTable("SFN_FATURA"), "DATAVENCIMENTO") And SALDO < 0) Then
      		bsShowMessage("Usuário sem permissão para alterar a data de vencimento da fatura.", "I")
      		Exit Sub
    	End If
	End If

End Sub

Public Sub BOTAOMODIFICARCODIGOFOLHA_OnClick()

  SessionVar("HFATURA") = CurrentQuery.FieldByName("HANDLE").AsString

  If (VerficaCodigoFolha = "") Then
	Dim INTERFACE002 As Object
	Dim vsMensagem As String
	Dim vcContainer As CSDContainer

	Set INTERFACE002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

	INTERFACE002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0094", _
					   "Modificar Código Folha",  _
					   0, _
					   150, _
					   300, _
					   False, _
					   vsMensagem, _
					   vcContainer)

	Set INTERFACE002 = Nothing
  Else
    bsShowMessage(VerficaCodigoFolha, "I")
  End If

End Sub

Public Sub TABLE_AfterScroll()
  Dim CPFCGC As String
  Dim mascara As String
  Dim IntBenef As Object
  Dim ContaFin As Object
  Set ContaFin = NewQuery
  ContaFin.Clear
  ContaFin.Add("SELECT TABRESPONSAVEL, BENEFICIARIO, PRESTADOR, PESSOA")
  ContaFin.Add("FROM SFN_CONTAFIN WHERE HANDLE = :HANDLECONTAFIN")
  ContaFin.ParamByName("HANDLECONTAFIN").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  ContaFin.Active = True
  If ContaFin.EOF Then
    ROTULORESPONSAVEL.Text = "*** CONTA FINANCEIRA NÃO ENCONTRADA ***"
  Else
    If ContaFin.FieldByName("TABRESPONSAVEL").Value = 1 Then
      Dim Beneficiario As Object
      Set Beneficiario = NewQuery
      Beneficiario.Clear
      Beneficiario.Add("SELECT BENEFICIARIO, NOME FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLEBENEFICIARIO")
      Beneficiario.ParamByName("HANDLEBENEFICIARIO").Value = ContaFin.FieldByName("BENEFICIARIO").AsInteger
      Beneficiario.Active = True
      If Beneficiario.EOF Then
        ROTULORESPONSAVEL.Text = "*** BENEFICIÁRIO NÃO ENCONTRADO ***"
      Else
        Set IntBenef = CreateBennerObject("SamBeneficiario.Cadastro")
        mascara = ""
        IntBenef.Mascara(CurrentSystem, Beneficiario.FieldByName("BENEFICIARIO").AsString, "", mascara)
        ROTULORESPONSAVEL.Text = "BENEFICIÁRIO: " + mascara + " - " + Beneficiario.FieldByName("NOME").Value
        ' ROTULORESPONSAVEL.Text ="BENEFICIÁRIO: " + _
        '    Format(Beneficiario.FieldByName("BENEFICIARIO").Value,"000000\.000000\.00")+ _
        '    " - " +Beneficiario.FieldByName("NOME").Value
      End If
      Beneficiario.Active = False
      Set Beneficiario = Nothing
    Else
      If ContaFin.FieldByName("TABRESPONSAVEL").Value = 2 Then
        Dim Prestador As Object
        Set Prestador = NewQuery
        Prestador.Clear
        Prestador.Add("SELECT PRESTADOR, NOME, CPFCNPJ FROM SAM_PRESTADOR WHERE HANDLE = :HANDLEPRESTADOR")
        Prestador.ParamByName("HANDLEPRESTADOR").Value = ContaFin.FieldByName("PRESTADOR").AsInteger
        Prestador.Active = True
        If Prestador.EOF Then
          ROTULORESPONSAVEL.Text = "*** PRESTADOR NÃO ENCONTRADO ***"
        Else
          ROTULORESPONSAVEL.Text = "PRESTADOR: "
          ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + Prestador.FieldByName("PRESTADOR").AsString

          If Len(Prestador.FieldByName("CPFCNPJ").AsString)<= 11 Then
            ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + " - " + Format(Prestador.FieldByName("CPFCNPJ").AsString, "000\.000\.000\-00")
          Else
            ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + " - " + Format(Prestador.FieldByName("CPFCNPJ").AsString, "00\.000\.000\/0000\-00")
          End If

          ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + " - " + Prestador.FieldByName("NOME").AsString
        End If
        Prestador.Active = False
        Set Prestador = Nothing
      Else
        If ContaFin.FieldByName("TABRESPONSAVEL").Value = 3 Then
          Dim Pessoa As Object
          Set Pessoa = NewQuery
          Pessoa.Clear
          Pessoa.Add("SELECT CNPJCPF, NOME FROM SFN_PESSOA WHERE HANDLE = :HANDLEPESSOA")
          Pessoa.ParamByName("HANDLEPESSOA").Value = ContaFin.FieldByName("PESSOA").AsInteger
          Pessoa.Active = True
          If Pessoa.EOF Then
            ROTULORESPONSAVEL.Text = "*** PESSOA NÃO ENCONTRADA ***"
          Else
            ROTULORESPONSAVEL.Text = "PESSOA: "
            If Len(Pessoa.FieldByName("CNPJCPF").AsString)<= 11 Then
              ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + _
                                       Format(Pessoa.FieldByName("CNPJCPF").AsString, "000\.000\.000\-00")
            Else
              ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + _
                                       Format(Pessoa.FieldByName("CNPJCPF").AsString, "00\.000\.000\/0000\-00")
            End If
            ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + " - " + _
                                     Pessoa.FieldByName("NOME").AsString
          End If
          Pessoa.Active = False
          Set Pessoa = Nothing
        End If
      End If
    End If
  End If
  ContaFin.Active = False
  Set IntBenef = Nothing
  Set ContaFin = Nothing

  BOTAOCANCELAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "A")
  BOTAOESTORNOCANCELA.Enabled = (CurrentQuery.FieldByName("CANCVALOR").AsString <> "") And (CurrentQuery.FieldByName("CANCVALOR").AsFloat > 0)
  BOTAOBAIXAR.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "A")
  BOTAOESTORNOBAIXA.Enabled = (CurrentQuery.FieldByName("SITUACAO").AsString = "B")
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
    SessionVar("HFATURA") = CurrentQuery.FieldByName("HANDLE").AsString
	Select Case CommandID
		Case "BOTAOATUALIZAR"
			BOTAOATUALIZAR_OnClick
		Case "BOTAOCONTAFINANCEIRA"
			BOTAOCONTAFINANCEIRA_OnClick
		Case "BOTAOEXCLUIR"
			BOTAOEXCLUIR_OnClick
		Case "BOTAOCOPIAR"
			BOTAOCOPIAR_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOESTORNOCANCELA"
			BOTAOESTORNOCANCELA_OnClick
		Case "BOTAOBAIXAR"
 		    BOTAOBAIXAR_OnClick
		Case "BOTAOESTORNOBAIXA"
			BOTAOESTORNOBAIXA_OnClick
		Case "BOTAOALTERARVENCIMENTO"
			BOTAOALTERARVENCIMENTO_OnClick
		Case "MODIFICARCODIGOFOLHA"
            BOTAOMODIFICARCODIGOFOLHA_OnClick
          If (VerficaCodigoFolha <> "") Then
            bsShowMessage(VerficaCodigoFolha, "E")
		    CanContinue = False
		  End If
	End Select
End Sub

Public Function VerficaCodigoFolha() As String

	Dim qVerificaCodigoFolha As BPesquisa
	Set qVerificaCodigoFolha = NewQuery

	qVerificaCodigoFolha.Add("SELECT F.SALDO SALDOFATURA,                          ")
	qVerificaCodigoFolha.Add("       C.BENEFICIARIO CONTAFINBENEFICIARIO,          ")
	qVerificaCodigoFolha.Add("       F.FOLHAPAGAMENTO FOLHAFATURA,                 ")
	qVerificaCodigoFolha.Add("       T.FOLHAPAGAMENTO FOLHACONTRATO                ")
	qVerificaCodigoFolha.Add("  FROM SFN_FATURA F                                  ")
	qVerificaCodigoFolha.Add("JOIN SFN_CONTAFIN C ON C.HANDLE = F.CONTAFINANCEIRA  ")
	qVerificaCodigoFolha.Add("JOIN SAM_BENEFICIARIO B ON B.HANDLE = C.BENEFICIARIO ")
	qVerificaCodigoFolha.Add("JOIN SAM_CONTRATO T ON T.HANDLE = B.CONTRATO         ")
	qVerificaCodigoFolha.Add("WHERE F.HANDLE = :HANDLEFATURA                       ")
	qVerificaCodigoFolha.ParamByName("HANDLEFATURA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	qVerificaCodigoFolha.Active = True

	VerficaCodigoFolha = ""

	If (qVerificaCodigoFolha.FieldByName("CONTAFINBENEFICIARIO").AsString = "") Then
      VerficaCodigoFolha =  "Não foi possível alterar o código folha da fatura. A fatura não pertence a uma conta financeira de beneficiário!"
      Exit Function
	End If

	If (qVerificaCodigoFolha.FieldByName("SALDOFATURA").AsInteger <= 0) Then
	  VerficaCodigoFolha = "Não foi possível alterar o código folha da fatura, o saldo da fatura deve ser maior que zero!"
	   Exit Function
	End If

	If (qVerificaCodigoFolha.FieldByName("FOLHAFATURA").AsString = "" Or qVerificaCodigoFolha.FieldByName("FOLHACONTRATO").AsString = "") Then
      VerficaCodigoFolha = "Não foi possível alterar o código folha da fatura. A fatura ou o contrato do beneficiário deve possuir dados de folha de pagamento!"
       Exit Function
	End If

End Function
