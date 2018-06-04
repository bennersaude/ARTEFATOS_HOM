'HASH: 827145C925EB9CAB726902499C749482
'Macro: SFN_CONTAFIN_ALTERACAO
' SFN_CONTAFIN-ALTERACAO
' Ultima alteraçao: 19/10/2000
'31/08/2000 16:30 Juliano
'03/07/2003 14:00 Celso - erros de parametro e sem transação

'#Uses "*CheckCPFCNPJ"
'#Uses "*bsShowMessage
'#Uses "*IsInt"

Option Explicit
Dim NaoGerarDocumentoAnterior As String
'André - SMS 24062 - 31/08/2004
Dim viBanco As Integer
Dim viAgencia As Integer
Dim VCC As String
Dim vDV As String
Dim Msg As String
Dim vBenef As String
Dim vOrdem As Integer
Dim vAlteraBanco As Boolean


Public Sub AGENCIA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SFN_AGENCIA.NOME|SFN_AGENCIA.AGENCIA"
  vCriterio = "SFN_AGENCIA.BANCO =" + CurrentQuery.FieldByName("BANCO").AsString
  vCampos = "Nome|Código"

  Dim textoAgencia As String
  textoAgencia = AGENCIA.Text
  If IsInt(TiraAcento(textoAgencia,True)) Then
    vOrdem = 2
  Else
    vOrdem = 1
  End If
  vHandle = interface.Exec(CurrentSystem, "SFN_AGENCIA", vColunas, vOrdem, vCampos, vCriterio, "Tabela de Agências", True, textoAgencia)

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("AGENCIA").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub BANCO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "NOME|CODIGO"

  vCampos = "Nome|Código"

 Dim textoBanco As String
  textoBanco = BANCO.Text
  If IsInt(TiraAcento(textoBanco,True)) Then
    vOrdem = 2
  Else
    vOrdem = 1
  End If

  vHandle = interface.Exec(CurrentSystem, "SFN_BANCO", vColunas, vOrdem, vCampos, vCriterio, "Tabela de Bancos", True, textoBanco)

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BANCO").Value = vHandle
    CurrentQuery.FieldByName("AGENCIA").Clear
  End If
  Set interface = Nothing
End Sub

'************* SMS 25916 - Kristian *******************

Public Sub BENEFICIARIODESTINOCOBRANCA_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.NOME"
  vCriterio = "SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL AND SAM_BENEFICIARIO.EHTITULAR = 'S'"
  vCriterio = vCriterio + " AND EXISTS (SELECT '1' FROM SAM_CONTRATO WHERE HANDLE = SAM_BENEFICIARIO.CONTRATO AND TABFOLHAPAGAMENTO = 2 AND TABPERMITERECEBERMENSALOUTROS = 2 AND LOCALFATURAMENTO = 'F' AND TIPOFATURAMENTO = (SELECT HANDLE FROM SIS_TIPOFATURAMENTO WHERE CODIGO = 130))"
  vCampos = "Beneficiario|Nome"

  vHandle = interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, 1, vCampos, vCriterio, "Tabela de Beneficiários", True, "")

  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BENEFICIARIODESTINOCOBRANCA").Value = vHandle
  End If
  Set interface = Nothing
End Sub 'FIM SMS 25916

Public Sub BOTAOCONFIRMA_OnClick()
  Dim VERIFICA As Object
  Dim CONTAFIN As Object
  Dim ALTERA As Object
  Dim CODBANCO As Object
  Dim CODAGENCIA As Object
  Dim CODOPERADORA As Object
  Set VERIFICA = NewQuery
  Set CONTAFIN = NewQuery
  Set ALTERA = NewQuery
  Set CODBANCO = NewQuery
  Set CODAGENCIA = NewQuery
  Set CODOPERADORA = NewQuery
  Dim vsDADOSANTERIORES As String
  Dim tipodoc As Object
  Set tipodoc = NewQuery
  Dim SQLCLASSECONTABIL As Object
  Set SQLCLASSECONTABIL = NewQuery
  Dim vClasseContabil As String

  If CurrentQuery.State <> 1 Then
     bsShowMessage("É necessário gravar as alterações antes de confirmar!", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("SITUACAO").AsString = "S" Then

    If (CurrentQuery.FieldByName("TABGERACAOREC").Value = 3) And (CurrentQuery.FieldByName("TIPODOCUMENTOREC").IsNull) Then
      'bsShowMessage("É necessário informar o tipo de documento para recebimento!", "I")
	  'Exit Sub

      If bsShowMessage("Tipo de documento para recebimento não informado!" + Chr(13) + "Deseja continuar ?", "Q") = vbNo Then
        Exit Sub
      End If

	End If

    If (CurrentQuery.FieldByName("TABGERACAOPAG").Value = 3) And (CurrentQuery.FieldByName("TIPODOCUMENTOPAG").IsNull) Then
	  'bsShowMessage("É necessário informar o tipo de documento para pagamento!", "I")
	  'Exit Sub
      If bsShowMessage("Tipo de documento para pagamento não informado!" + Chr(13) + "Deseja continuar ?", "Q") = vbNo Then
        Exit Sub
      End If

	End If

    'Seleciona o registro da Tabela SFN_CONTAFIN
    VERIFICA.Clear
    VERIFICA.Add("SELECT * FROM SFN_CONTAFIN")
    VERIFICA.Add("WHERE HANDLE=:HANDLE")
    VERIFICA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
    VERIFICA.Active = True

    If Not VERIFICA.FieldByName("CLASSECONTABIL").IsNull Then
      SQLCLASSECONTABIL.Clear
      SQLCLASSECONTABIL.Active = False
      SQLCLASSECONTABIL.Add("SELECT ESTRUTURA, DESCRICAO FROM SFN_CLASSECONTABIL WHERE HANDLE=:HANDLE")
      SQLCLASSECONTABIL.ParamByName("HANDLE").AsInteger = VERIFICA.FieldByName("CLASSECONTABIL").AsInteger
      SQLCLASSECONTABIL.Active = True

      vClasseContabil = SQLCLASSECONTABIL.FieldByName("ESTRUTURA").AsString + " - " + SQLCLASSECONTABIL.FieldByName("DESCRICAO").AsString
    Else
      vClasseContabil = ""
    End If

    vsDADOSANTERIORES = ""

    'Testa se a alteraçao é do tipo Conta corrente
    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 1 Then
      vsDADOSANTERIORES = "Geração pagamento  = Conta corrente"
    End If

    If VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 1 Then
      If vsDADOSANTERIORES <> "" Then
        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
                          "Geração recebimento = Conta corrente"
      Else
        vsDADOSANTERIORES = "Geração recebimento = Conta corrente"
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 1 Or VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 1Then

      'Busca o número do Banco
      CODBANCO.Clear
      CODBANCO.Add("SELECT CODIGO FROM SFN_BANCO WHERE HANDLE = :PBANCO")
      CODBANCO.ParamByName("PBANCO").AsInteger = VERIFICA.FieldByName("BANCO").AsInteger
      CODBANCO.Active = True

      'Busca o número da Agência
      CODAGENCIA.Clear
      CODAGENCIA.Add("SELECT AGENCIA FROM SFN_AGENCIA WHERE HANDLE = :PAGENCIA")
      CODAGENCIA.ParamByName("PAGENCIA").Value = VERIFICA.FieldByName("AGENCIA").AsInteger
      CODAGENCIA.Active = True

      'Registra os dados anteriores na tabela de alteraçoes
      vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
                        "Banco = " + CODBANCO.FieldByName("CODIGO").AsString + " | " + _
                        "Agência = " + CODAGENCIA.FieldByName("AGENCIA").AsString + " | " + _
                        "Conta Corrente = " + VERIFICA.FieldByName("CONTACORRENTE").AsString + " | " + _
                        "DV = " + VERIFICA.FieldByName("DV").AsString + " | " + _
                        "Nome = " + VERIFICA.FieldByName("CCNOME").AsString + " | " + _
                        "CPF = " + VERIFICA.FieldByName("CCCPFCNPJ").AsString + " | " + _
                        "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                        "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
                        "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString


    End If


    'Testa se a alteração é do tipo cartao de credito
    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 2 Then
      If vsDADOSANTERIORES <> "" Then
        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
                          "Geração pagamento  = Cartão de crédito"
      Else
        vsDADOSANTERIORES = "Geração pagamento  = Cartão de crédito"
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 2 Then
      If vsDADOSANTERIORES <> "" Then
        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
                          "Geração recebimento  = Cartão de crédito"
      Else
        vsDADOSANTERIORES = "Geração recebimento  = Cartão de crédito"
      End If
    End If

    If VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 2 Or VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 2 Then

      'vERIFICA O NOME DA OPERADORA DO CARTAO
      CODOPERADORA.Clear
      CODOPERADORA.Add("SELECT CODIGO, NOME")
      CODOPERADORA.Add("  FROM SFN_CARTAOOPERADORA")
      CODOPERADORA.Add(" WHERE HANDLE = :HANDLE")
      CODOPERADORA.ParamByName("HANDLE").Value = VERIFICA.FieldByName("CARTAOOPERADORA").AsInteger
      CODOPERADORA.Active = True

      vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
                        "Código operadora = " + CODOPERADORA.FieldByName("CODIGO").AsString + " | " + _
                        "Nome operadora = " + CODOPERADORA.FieldByName("NOME").AsString + " | " + _
                        "Número do cartão = " + VERIFICA.FieldByName("CARTAOCREDITO").AsString + " | " + _
                        "Validade = " + VERIFICA.FieldByName("CARTAOVALIDADE").AsString + " | " + _
                        "Não gerar documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
                        "Não cobrar tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
                        "Não parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString

    End If

    'Testa se a alteração é do tipo Título
    If (VERIFICA.FieldByName("TABGERACAOPAG").AsInteger = 3) Then

      tipodoc.Active = False
      tipodoc.Clear
      tipodoc.Add("SELECT DESCRICAO FROM SFN_TIPODOCUMENTO")
      tipodoc.Add(" WHERE HANDLE = :HANDLEDOC")
      tipodoc.ParamByName("HANDLEDOC").Value = VERIFICA.FieldByName("TIPODOCUMENTOPAG").AsInteger
      tipodoc.Active = True

	  If Not (VERIFICA.FieldByName("TIPODOCUMENTOPAG").IsNull) Then

	      If vsDADOSANTERIORES <> "" Then
	        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
	                          "Geração pagamento  = Título" + Chr(13) + _
	                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      Else
	        vsDADOSANTERIORES = "Geração pagamento  = Título" + Chr(13) + _
	                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      End If
	  Else

		  If vsDADOSANTERIORES <> "" Then
	        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
	                          "Geração pagamento  = Título" + Chr(13) + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      Else
	        vsDADOSANTERIORES = "Geração pagamento  = Título" + Chr(13) + _
	                            "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                            "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      End If
      End If
    End If

    If (VERIFICA.FieldByName("TABGERACAOREC").AsInteger = 3) Then

      tipodoc.Active = False
      tipodoc.Clear
      tipodoc.Add("SELECT DESCRICAO FROM SFN_TIPODOCUMENTO")
      tipodoc.Add(" WHERE HANDLE = :HANDLEDOC")
      tipodoc.ParamByName("HANDLEDOC").Value = VERIFICA.FieldByName("TIPODOCUMENTOREC").AsInteger
      tipodoc.Active = True

	  If Not (VERIFICA.FieldByName("TIPODOCUMENTOREC").IsNull) Then

	      If vsDADOSANTERIORES <> "" Then
	        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
	                          "Geração recebimento = Título" + Chr(13) + _
	                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      Else
	        vsDADOSANTERIORES = "Geração recebimento = Título" + Chr(13) + _
	                          "Tipo do documento = " + tipodoc.FieldByName("DESCRICAO").AsString + " | " + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      End If
	  Else

		  If vsDADOSANTERIORES <> "" Then
	        vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + _
	                          "Geração recebimento = Título" + Chr(13) + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      Else
	        vsDADOSANTERIORES = "Geração recebimento = Título" + Chr(13) + _
	                          "Não Gerar Documento = " + VERIFICA.FieldByName("NAOGERARDOCUMENTO").AsString + " | " + _
	                          "Não Cobrar Tarifa = " + VERIFICA.FieldByName("NAOCOBRARTARIFA").AsString + " | " + _
	                          "Não Parcelar = " + VERIFICA.FieldByName("NAOPARCELAR").AsString
	      End If
      End If
    End If

    On Error GoTo ERRO

    StartTransaction

    If vClasseContabil <> "" Then
      vsDADOSANTERIORES = vsDADOSANTERIORES + Chr(13) + "Classe Contábil = " + vClasseContabil
    Else
      'DADOSANTERIORES = DADOSANTERIORES + Chr(13) + "Classe Contábil = " + vC
    End If

    'Registra os dados anteriores na tabela de alteraçoes
    ALTERA.Clear
    ALTERA.Add("UPDATE SFN_CONTAFIN_ALTERACAO Set DADOSANTERIORES = :PDADOS WHERE HANDLE = :PHANDLE")
    ALTERA.ParamByName("PDADOS").AsMemo = vsDADOSANTERIORES
    ALTERA.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    ALTERA.ExecSQL

    If CurrentQuery.FieldByName("TABCOBRANCAOUTRORESPONSAVEL").AsInteger = 2 Then
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN ")
      CONTAFIN.Add("   SET TABCOBRANCAOUTRORESPONSAVEL = :TABRESP,")
      CONTAFIN.Add("       BENEFICIARIODESTINOCOBRANCA = :BENEF   ")
      CONTAFIN.Add(" WHERE HANDLE = :PHANDLE")

      CONTAFIN.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      CONTAFIN.ParamByName("TABRESP").AsInteger = CurrentQuery.FieldByName("TABCOBRANCAOUTRORESPONSAVEL").AsInteger
      CONTAFIN.ParamByName("BENEF").AsInteger = CurrentQuery.FieldByName("BENEFICIARIODESTINOCOBRANCA").AsInteger

      CONTAFIN.ExecSQL
    End If

      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET               				")

      CONTAFIN.Add(" HANDLE = HANDLE,                                   ")

      If CurrentQuery.FieldByName("BANCO").AsInteger > 0 Then
        CONTAFIN.Add("                        BANCO = :PBANCO,		    ")
        CONTAFIN.ParamByName("PBANCO").Value = CurrentQuery.FieldByName("BANCO").AsInteger
      Else
        CONTAFIN.Add("                        BANCO = NULL,		    ")
      End If

      If CurrentQuery.FieldByName("AGENCIA").AsInteger > 0 Then
        CONTAFIN.Add("                        AGENCIA = :PAGENCIA,		")
        CONTAFIN.ParamByName("PAGENCIA").Value = CurrentQuery.FieldByName("AGENCIA").AsInteger
      Else
        CONTAFIN.Add("                        AGENCIA = NULL,		")
      End If

      CONTAFIN.Add("                        CONTACORRENTE = :PCONTA,	")
      CONTAFIN.ParamByName("PCONTA").Value = CurrentQuery.FieldByName("CONTACORRENTE").AsString

      CONTAFIN.Add("                        DV = :PDV,                  ")
      CONTAFIN.ParamByName("PDV").Value = CurrentQuery.FieldByName("DV").AsString

      CONTAFIN.Add("                        CCNOME = :PCCNOME,          ")
      CONTAFIN.ParamByName("PCCNOME").Value = CurrentQuery.FieldByName("CCNOME").AsString

      CONTAFIN.Add("                        CCCPFCNPJ = :PCCCPFCNPJ,      ")
      CONTAFIN.ParamByName("PCCCPFCNPJ").Value = CurrentQuery.FieldByName("CCCPFCNPJ").AsString

      CONTAFIN.Add("                        NAOPARCELAR = :NAOPARCELAR      ")
      CONTAFIN.ParamByName("NAOPARCELAR").Value = CurrentQuery.FieldByName("NAOPARCELAR").AsString

      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      CONTAFIN.ExecSQL

    If CurrentQuery.FieldByName("TABGERACAOPAG").AsInteger = 1 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOPAG = 1,				")
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC,      ")
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA,   ")
        CONTAFIN.Add("                        CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("PDOC").Value = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = CurrentQuery.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      CONTAFIN.ExecSQL

    End If

    If CurrentQuery.FieldByName("TABGERACAOREC").AsInteger = 1 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOREC = 1,              ")
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC,      ")
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA,     ")
        CONTAFIN.Add("                        CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                        NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("PDOC").Value = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = CurrentQuery.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If


    If CurrentQuery.FieldByName("TABGERACAOPAG").AsInteger = 2 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOPAG = 2, CARTAOOPERADORA = :OPERADORA, CARTAOCREDITO = :PCARTAO,")
      CONTAFIN.Add("                        CARTAOVALIDADE = :PVALIDADE, NAOGERARDOCUMENTO = :PDOC,                   ")
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("OPERADORA").Value = CurrentQuery.FieldByName("CARTAOOPERADORA").AsInteger
      CONTAFIN.ParamByName("PCARTAO").Value = CurrentQuery.FieldByName("CARTAOCREDITO").AsString
      CONTAFIN.ParamByName("PVALIDADE").Value = CurrentQuery.FieldByName("CARTAOVALIDADE").AsDateTime
      CONTAFIN.ParamByName("PDOC").Value = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = CurrentQuery.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If

    If CurrentQuery.FieldByName("TABGERACAOREC").AsInteger = 2 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOREC = 2, CARTAOOPERADORA = :OPERADORA, CARTAOCREDITO = :PCARTAO,")
      CONTAFIN.Add("                        CARTAOVALIDADE = :PVALIDADE, NAOGERARDOCUMENTO = :PDOC, ")
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      CONTAFIN.ParamByName("OPERADORA").Value = CurrentQuery.FieldByName("CARTAOOPERADORA").AsInteger
      CONTAFIN.ParamByName("PCARTAO").Value = CurrentQuery.FieldByName("CARTAOCREDITO").AsString
      CONTAFIN.ParamByName("PVALIDADE").Value = CurrentQuery.FieldByName("CARTAOVALIDADE").AsDateTime
      CONTAFIN.ParamByName("PDOC").Value = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = CurrentQuery.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If


    If CurrentQuery.FieldByName("TABGERACAOPAG").AsInteger = 3 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOPAG = 3,")
      If Not CurrentQuery.FieldByName("TIPODOCUMENTOPAG").IsNull Then
          CONTAFIN.Add(" TIPODOCUMENTOPAG = :tipodoc,")
      Else
          CONTAFIN.Add(" TIPODOCUMENTOPAG = NULL,")
      End If
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC, ")
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      If Not CurrentQuery.FieldByName("TIPODOCUMENTOPAG").IsNull Then
        CONTAFIN.ParamByName("TIPODOC").Value = CurrentQuery.FieldByName("TIPODOCUMENTOPAG").AsInteger
      End If
      CONTAFIN.ParamByName("PDOC").Value = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = CurrentQuery.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If

    If CurrentQuery.FieldByName("TABGERACAOREC").AsInteger = 3 Then

      'Atualiza as alterações na tabela conta financeira
      CONTAFIN.Clear
      CONTAFIN.Add("UPDATE SFN_CONTAFIN SET TABGERACAOREC = 3, ")
      If CurrentQuery.FieldByName("TIPODOCUMENTOREC").AsInteger > 0 Then
         CONTAFIN.Add(" TIPODOCUMENTOREC = :tipodoc,")
      Else
         CONTAFIN.Add(" TIPODOCUMENTOREC = null,")
      End If
      CONTAFIN.Add("                        NAOGERARDOCUMENTO = :PDOC, ")
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA, CLASSECONTABIL = :CLASSECONTABIL")
      Else
        CONTAFIN.Add("                      NAOCOBRARTARIFA = :PTARIFA,")
        CONTAFIN.Add("                        CLASSECONTABIL = NULL")
      End If
      CONTAFIN.Add("WHERE HANDLE = :PHANDLE")
      If CurrentQuery.FieldByName("TIPODOCUMENTOREC").AsInteger > 0 Then
         CONTAFIN.ParamByName("TIPODOC").Value = CurrentQuery.FieldByName("TIPODOCUMENTOREC").AsInteger
      End If
      CONTAFIN.ParamByName("PDOC").Value = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
      CONTAFIN.ParamByName("PTARIFA").Value = CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString
      CONTAFIN.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      If Not CurrentQuery.FieldByName("CLASSECONTABIL").IsNull Then
        CONTAFIN.ParamByName("CLASSECONTABIL").Value = CurrentQuery.FieldByName("CLASSECONTABIL").Value
      End If
      CONTAFIN.ExecSQL
    End If

	' SMS 76561 Danilo Raisi - 07/02/2007
	'Inclui na SFN_CONTAFIN o registro da última alteração efetuada
	ALTERA.Clear
	ALTERA.Add("UPDATE SFN_CONTAFIN 					    ")
	ALTERA.Add("   SET ALTERACAODATA    = :ALTERACAODATA,   ")
	ALTERA.Add("	   ALTERACAOUSUARIO = :ALTERACAOUSUARIO ")
	ALTERA.Add(" WHERE HANDLE = (:PHANDLE)					")
	ALTERA.ParamByName("ALTERACAODATA").AsDateTime = ServerNow
	ALTERA.ParamByName("ALTERACAOUSUARIO").AsInteger = CurrentUser
	ALTERA.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
	ALTERA.ExecSQL
	' Fim SMS 76561 Danilo Raisi - 07/02/2007

    'Grava a data e o usuário responsável pela confirmação
    ALTERA.Clear
    ALTERA.Add("UPDATE SFN_CONTAFIN_ALTERACAO SET SITUACAO = 'C', CONFIRMADODATA = :PDATA, CONFIRMADOUSUARIO = :PUSUARIO")
    ALTERA.Add("WHERE HANDLE = :PHANDLE")
    ALTERA.ParamByName("PDATA").Value = ServerDate
    ALTERA.ParamByName("PUSUARIO").Value = CurrentUser
    ALTERA.ParamByName("PHANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    ALTERA.ExecSQL

	Dim tabelaIntegracao As String
	Dim SQL As BPesquisa

	Set SQL = NewQuery

	tabelaIntegracao = ""
    If (CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1) Then
        tabelaIntegracao = "SAM_BENEFICIARIO"
        SQL.Active = False
      	SQL.Clear
      	SQL.Add("SELECT BENEFICIARIO HANDLEORIGEM FROM SFN_CONTAFIN WHERE HANDLE = :CONTAFINALTERACAO")
      	SQL.ParamByName("CONTAFINALTERACAO").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      	SQL.Active = True
    ElseIf (CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 2) Then
        SQL.Active = False
      	SQL.Clear
      	SQL.Add("SELECT CFIN.PRESTADOR HANDLEORIGEM, PREST.RECEBEDOR FROM SFN_CONTAFIN CFIN JOIN SAM_PRESTADOR PREST ON PREST.HANDLE = CFIN.PRESTADOR WHERE CFIN.HANDLE = :CONTAFINALTERACAO")
      	SQL.ParamByName("CONTAFINALTERACAO").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      	SQL.Active = True

      	If SQL.FieldByName("RECEBEDOR").AsString = "S" Then
      	  tabelaIntegracao = "SAM_PRESTADOR"
 		End If
    ElseIf (CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 3) Then
		tabelaIntegracao = "SFN_PESSOA"
      	SQL.Clear
      	SQL.Add("SELECT PESSOA HANDLEORIGEM FROM SFN_CONTAFIN WHERE HANDLE = :CONTAFINALTERACAO")
      	SQL.ParamByName("CONTAFINALTERACAO").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      	SQL.Active = True
    End If

	If tabelaIntegracao <> "" Then
	  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
	  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

	  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, SQL.FieldByName("HANDLEORIGEM").AsInteger)
	  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, tabelaIntegracao)
	  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

	  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
	End If

	Set SQL = Nothing

    Commit

    CurrentQuery.Active = False
    CurrentQuery.Active = True

    bsShowMessage("Conta Financeira alterada!", "I")

  Else
    bsShowMessage("Conta Financeira já está confirmada!", "I")

  End If

  GoTo FINALIZA
ERRO:
  Rollback
  bsShowMessage (Error, "I")

FINALIZA:
  Set VERIFICA = Nothing
  Set CONTAFIN = Nothing
  Set ALTERA = Nothing
  Set CODBANCO = Nothing
  Set CODAGENCIA = Nothing
  Set tipodoc = Nothing

End Sub

Public Sub CCCPFCNPJ_OnExit()

  If CurrentQuery.FieldByName("CCCPFCNPJ").AsString <> "" Then
    If Not CheckCPFCNPJ(RetornaNumeros(CurrentQuery.FieldByName("CCCPFCNPJ").AsString), 0, True, Msg) Then
      bsShowMessage(Msg, "I")
      CCCPFCNPJ.SetFocus
    End If
  End If
End Sub

Public Sub TABCOBRANCAOUTRORESPONSAVEL_OnChanging(AllowChange As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").Value = "C" Then
    bsShowMessage("Esta alteração já está confirmada", "I")
    AllowChange = False
    Exit Sub
  End If
End Sub

Public Sub TABGERACAOPAG_OnChange()

  If (vAlteraBanco) Then
    If (TABGERACAOPAG.PageIndex = 0) Then 'Conta Corrente
      Dim SQL2 As Object
      Set SQL2 = NewQuery

      SQL2.Clear
      SQL2.Active = False
      SQL2.Add("SELECT T.BANCO ")
      SQL2.Add("FROM SFN_CONTAFIN CF, SAM_BENEFICIARIO B,")
      SQL2.Add("     SAM_CONTRATO C, SFN_TESOURARIA T")
      SQL2.Add("WHERE CF.BENEFICIARIO=B.HANDLE AND B.CONTRATO=C.HANDLE")
      SQL2.Add("      AND C.TESOURARIARECEBIMENTO=T.HANDLE")
      SQL2.Add("      AND CF.HANDLE=:CONTAFIN AND T.TABTIPO=1")
      SQL2.ParamByName("CONTAFIN").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      SQL2.Active = True

      If Not SQL2.EOF Then 'Caso encontrar o banco
        AlterarBancoTesouraria(SQL2.FieldByName("BANCO").AsInteger)
      End If
      Set SQL2 = Nothing
    End If
  End If
End Sub

Public Sub TABGERACAOPAG_OnChanging(AllowChange As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").Value = "C" Then
    bsShowMessage("Esta alteração já está confirmada", "I")
    AllowChange = False
    Exit Sub
  End If

End Sub

Public Sub TABGERACAOREC_OnChange()

  If (vAlteraBanco) Then
    If (TABGERACAOREC.PageIndex = 0) Then 'Conta Corrente
      Dim SQL2 As Object
      Set SQL2 = NewQuery

      SQL2.Clear
      SQL2.Active = False
      SQL2.Add("SELECT T.BANCO ")
      SQL2.Add("FROM SFN_CONTAFIN CF, SAM_BENEFICIARIO B,")
      SQL2.Add("     SAM_CONTRATO C, SFN_TESOURARIA T")
      SQL2.Add("WHERE CF.BENEFICIARIO=B.HANDLE AND B.CONTRATO=C.HANDLE")
      SQL2.Add("      AND C.TESOURARIARECEBIMENTO=T.HANDLE")
      SQL2.Add("      AND CF.HANDLE=:CONTAFIN AND T.TABTIPO=1")
      SQL2.ParamByName("CONTAFIN").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
      SQL2.Active = True

      If Not SQL2.EOF Then 'Caso encontrar o banco
        AlterarBancoTesouraria(SQL2.FieldByName("BANCO").AsInteger)
      End If
      Set SQL2 = Nothing
    End If
  End If
End Sub

Public Sub TABGERACAOREC_OnChanging(AllowChange As Boolean)
  If CurrentQuery.FieldByName("SITUACAO").Value = "C" Then
    bsShowMessage("Esta alteração já está confirmada", "I")
    AllowChange = False
    Exit Sub
  End If

  'Alteração Henrique 15/10/2002  SMS 13097
  'If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
  '  CCNOME.Visible = True
  '  CCCPFCNPJ.Visible = True
  'Else
  '  CCNOME.Visible = False
  '  CCCPFCNPJ.Visible = False
  'End If
End Sub

Public Sub TABLE_AfterInsert()
  Dim SQL As Object
  Set SQL = NewQuery
  vAlteraBanco = True

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT CF.*, BE.NOME NOMEBENEFICIARIO,")
  SQL.Add("       PR.NOME NOMEPRESTADOR, PE.NOME NOMEPESSOA")
  SQL.Add("FROM SFN_CONTAFIN CF")
  SQL.Add("     LEFT JOIN SAM_BENEFICIARIO BE ON (BE.HANDLE = CF.BENEFICIARIO)")
  SQL.Add("     LEFT JOIN SAM_PRESTADOR PR ON (PR.HANDLE = CF.PRESTADOR)")
  SQL.Add("     LEFT JOIN SFN_PESSOA PE ON (PE.HANDLE = CF.PESSOA)")
  SQL.Add("WHERE CF.HANDLE = :HANDLE")

  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  SQL.Active = True

  CurrentQuery.FieldByName("TABGERACAOPAG").AsInteger = SQL.FieldByName("TABGERACAOPAG").AsInteger
  CurrentQuery.FieldByName("TABGERACAOREC").AsInteger = SQL.FieldByName("TABGERACAOREC").AsInteger
  CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = SQL.FieldByName("TABRESPONSAVEL").AsInteger
  CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString = SQL.FieldByName("NAOGERARDOCUMENTO").AsString
  CurrentQuery.FieldByName("NAOCOBRARTARIFA").AsString = SQL.FieldByName("NAOCOBRARTARIFA").AsString
  CurrentQuery.FieldByName("NAOPARCELAR").AsString = SQL.FieldByName("NAOPARCELAR").AsString
  CurrentQuery.FieldByName("CLASSECONTABIL").Value = SQL.FieldByName("CLASSECONTABIL").Value
  CurrentQuery.FieldByName("TIPODOCUMENTOREC").Value = SQL.FieldByName("TIPODOCUMENTOREC").Value
  CurrentQuery.FieldByName("TIPODOCUMENTOPAG").Value = SQL.FieldByName("TIPODOCUMENTOPAG").Value

  If (SQL.FieldByName("TABGERACAOPAG").AsInteger = 1) Or (SQL.FieldByName("TABGERACAOREC").AsInteger = 1) Then
    CurrentQuery.FieldByName("BANCO").Value = SQL.FieldByName("BANCO").Value
    CurrentQuery.FieldByName("AGENCIA").Value = SQL.FieldByName("AGENCIA").Value
    CurrentQuery.FieldByName("CONTACORRENTE").AsString = SQL.FieldByName("CONTACORRENTE").AsString
    CurrentQuery.FieldByName("DV").Value = SQL.FieldByName("DV").AsString
    CurrentQuery.FieldByName("CCNOME").Value = SQL.FieldByName("CCNOME").AsString
    CurrentQuery.FieldByName("CCCPFCNPJ").Value = SQL.FieldByName("CCCPFCNPJ").AsString
  End If

  If (SQL.FieldByName("TABGERACAOPAG").AsInteger = 2) Or (SQL.FieldByName("TABGERACAOREC").AsInteger = 2) Then
    CurrentQuery.FieldByName("CARTAOOPERADORA").AsString = SQL.FieldByName("CARTAOOPERADORA").AsString
    CurrentQuery.FieldByName("CARTAOCREDITO").AsString = SQL.FieldByName("CARTAOCREDITO").AsString
    CurrentQuery.FieldByName("CARTAOVALIDADE").Value = SQL.FieldByName("CARTAOVALIDADE").AsDateTime
  End If


  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
    ROTULORESPONSAVEL.Text = "BENEFICIARIO - " + SQL.FieldByName("NOMEBENEFICIARIO").AsString

    If (SQL.FieldByName("TABGERACAOPAG").AsInteger = 1) Or (SQL.FieldByName("TABGERACAOREC").AsInteger = 1) Then
      Dim SQL2 As Object
      Set SQL2 = NewQuery

      SQL2.Clear
      SQL2.Active = False
      SQL2.Add("SELECT T.BANCO ")
      SQL2.Add("FROM SAM_BENEFICIARIO B, SAM_CONTRATO C,")
      SQL2.Add("     SFN_TESOURARIA T")
      SQL2.Add("WHERE B.CONTRATO=C.HANDLE")
      SQL2.Add("      AND C.TESOURARIARECEBIMENTO=T.HANDLE ")
      SQL2.Add("      AND B.HANDLE=:BENEF AND T.TABTIPO=1")
      SQL2.ParamByName("BENEF").AsInteger = SQL.FieldByName("BENEFICIARIO").AsInteger
      SQL2.Active = True

      If Not SQL2.EOF Then 'Caso encontrar o banco
        If CurrentQuery.FieldByName("BANCO").IsNull Then
          CurrentQuery.FieldByName("BANCO").AsInteger = SQL2.FieldByName("BANCO").AsInteger
        ElseIf CurrentQuery.FieldByName("BANCO").AsInteger <> SQL2.FieldByName("BANCO").AsInteger Then
		  If WebMode Then
		  	If(CurrentEntity.TransitoryVars("ALTERARBANCOCONTAFINPORTESOURARIARECEBEDORA").IsPresent)Then
  				AlterarBancoTesouraria(CurrentEntity.TransitoryVars("ALTERARBANCOCONTAFINPORTESOURARIARECEBEDORA").AsInteger)
  			End If

          ElseIf bsShowMessage("Banco da Tesouraria do Contrato é diferente do Banco atual." + Chr(13) + "Deseja alterar ?", "Q") = vbYes Then
            AlterarBancoTesouraria(SQL2.FieldByName("BANCO").AsInteger)
		  Else
		    vAlteraBanco = False
          End If

        End If

      End If
      Set SQL2 = Nothing
    End If
  End If

  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then
    ROTULORESPONSAVEL.Text = "PRESTADOR - " + SQL.FieldByName("NOMEPRESTADOR").AsString
  End If

  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 3 Then
    ROTULORESPONSAVEL.Text = "PESSOA - " + SQL.FieldByName("NOMEPESSOA").AsString
  End If


  Set SQL = Nothing
  CurrentQuery.FieldByName("TABGERACAOREC").Value = 1
  CurrentQuery.FieldByName("TABGERACAOPAG").Value = 1
  CurrentQuery.FieldByName("TABCOBRANCAOUTRORESPONSAVEL").Value = 1

End Sub

Public Sub TABLE_AfterPost()
  vAlteraBanco = False
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  SQL.Add("SELECT CF.HANDLE,")
  SQL.Add("       CF.TABRESPONSAVEL,")
  SQL.Add("       BE.NOME NOMEBENEFICIARIO,")
  SQL.Add("       PR.NOME NOMEPRESTADOR,")
  SQL.Add("       PE.NOME NOMEPESSOA,")
  SQL.Add("       BE.MATRICULA, BE.HANDLE BENEFICIARIO")
  SQL.Add("FROM SFN_CONTAFIN CF")
  SQL.Add("     LEFT JOIN SAM_BENEFICIARIO BE ON (BE.HANDLE = CF.BENEFICIARIO)")
  SQL.Add("     LEFT JOIN SAM_PRESTADOR PR ON (PR.HANDLE = CF.PRESTADOR)")
  SQL.Add("     LEFT JOIN SFN_PESSOA PE ON (PE.HANDLE = CF.PESSOA)")
  SQL.Add("WHERE CF.HANDLE = :PHANDLE")


  SQL.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("CONTAFINANCEIRA").AsInteger
  SQL.Active = True

  If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
    ROTULORESPONSAVEL.Text = "BENEFICIARIO - " + SQL.FieldByName("NOMEBENEFICIARIO").AsString
  End If

  vBenef = SQL.FieldByName("BENEFICIARIO").AsString

  If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then
    ROTULORESPONSAVEL.Text = "PRESTADOR - " + SQL.FieldByName("NOMEPRESTADOR").AsString
  End If

  If SQL.FieldByName("TABRESPONSAVEL").AsInteger = 3 Then
    ROTULORESPONSAVEL.Text = "PESSOA - " + SQL.FieldByName("NOMEPESSOA").AsString
  End If
  Set SQL = Nothing

  If WebMode Then
  	'AGENCIA.WebLocalWhere = "SFN_AGENCIA.BANCO =" + CurrentQuery.FieldByName("BANCO").AsString
  	AGENCIA.WebLocalWhere = "A.BANCO = @CAMPO(BANCO)"
  	BENEFICIARIODESTINOCOBRANCA.WebLocalWhere = "A.DATACANCELAMENTO IS NULL AND A.EHTITULAR = 'S'"
  	BENEFICIARIODESTINOCOBRANCA.WebLocalWhere = BENEFICIARIODESTINOCOBRANCA.WebLocalWhere + " AND EXISTS (SELECT '1' FROM SAM_CONTRATO WHERE HANDLE = A.CONTRATO AND TABFOLHAPAGAMENTO = 2 AND TABPERMITERECEBERMENSALOUTROS = 2 AND LOCALFATURAMENTO = 'F' AND TIPOFATURAMENTO = (SELECT HANDLE FROM SIS_TIPOFATURAMENTO WHERE CODIGO = 130))"
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Texto As String
  Dim Doc As Object
  Set Doc = NewQuery
  Dim interface As Object
  Dim vRetorno As Boolean
  Dim vBanco As Integer, vAgencia As Long
  Dim vConta As String, vDV As String
  Dim CNPJCPF As String

  'SMS 25916 - Kristian
  Dim qBenef As Object
  Set qBenef = NewQuery

  If CurrentQuery.FieldByName("TABCOBRANCAOUTRORESPONSAVEL").AsInteger = 2 Then
    qBenef.Clear
    qBenef.Add("SELECT TABCOBRANCAOUTRORESPONSAVEL FROM SFN_CONTAFIN")
    qBenef.Add(" WHERE BENEFICIARIO = :BENEFICIARIO")
    qBenef.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIODESTINOCOBRANCA").AsInteger
    qBenef.Active = True
    If qBenef.FieldByName("TABCOBRANCAOUTRORESPONSAVEL").AsInteger = 2 Then
      bsShowMessage("Já existe outro responsável financeiro para este beneficiario que você escolheu!!!" + Chr(13) + _
             "Favor escolher outro beneficiário!!!" , "E")
      If VisibleMode Then
      	BENEFICIARIODESTINOCOBRANCA.SetFocus
      End If

      qBenef.Active = False
      Set qBenef = Nothing
      CanContinue = False
      Exit Sub
    End If
    qBenef.Active = False
    Set qBenef = Nothing
  End If
  'Fim SMS 25916

  If CurrentQuery.FieldByName("SITUACAO").Value = "C" Then
    bsShowMessage("Esta alteração já está confirmada.", "E")
    CanContinue = False
    Exit Sub
  End If

  'Keila SMS 9432
  'If (Not IsValidCPF(CurrentQuery.FieldByName("CCCPFCNPJ").AsString)) Or (Not IsValidCGC(CurrentQuery.FieldByName("CCCPFCNPJ").AsString)) Then
  If Not CheckCPFCNPJ(RetornaNumeros(CurrentQuery.FieldByName("CCCPFCNPJ").AsString), 0, True, Msg) Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If
  '  MsgBox "CPF/CNPJ Inválido!"
  '  CanContinue = False
  '  Exit Sub
  'End If


  'Juliano 12/03/01
  If CurrentQuery.FieldByName("TABGERACAOPAG").Value = 2 Then
    bsShowMessage("O tipo da geração de pagamento não pode ser cartão de crédito!", "E")
    CanContinue = False
    Exit Sub
  End If

  If VisibleMode Then

    Doc.Clear
    Doc.Add("Select * from SFN_Contafin where handle = :handle")
    Doc.ParamByName("handle").Value = CurrentQuery.FieldByName("contafinanceira").AsInteger
    Doc.Active = True

    NaoGerarDocumentoAnterior = Doc.FieldByName("NAOGERARDOCUMENTO").AsString

    Set Doc = Nothing

    If NaoGerarDocumentoAnterior <> CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString Then
      If CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString = "S" Then
        Texto = "A conta financeira foi modificada para NÃO GERAR documento ! "
      Else
        Texto = "A conta financeira foi modificada para GERAR documento ! "
     End If
      Texto = Texto + " Confirma ?"
      If bsShowMessage(Texto, "Q") <> vbYes Then
        CanContinue = False
        Exit Sub
      End If
    End If
  End If


  If CurrentQuery.FieldByName("TABGERACAOPAG").AsInteger = 1 Or CurrentQuery.FieldByName("TABGERACAOREC").AsInteger = 1 Then
    'sms 33910
    Dim qaux As Object
    Set qaux = NewQuery
    qaux.Clear
    qaux.Add("SELECT PERMITEAGENCIASEMDV FROM SFN_BANCO WHERE HANDLE = :HANDLE")
    qaux.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BANCO").AsInteger
    qaux.Active = True
    If (CurrentQuery.FieldByName("BANCO").IsNull) Or (CurrentQuery.FieldByName("AGENCIA").IsNull) Or _
       (CurrentQuery.FieldByName("CONTACORRENTE").IsNull) Or (CurrentQuery.FieldByName("DV").IsNull And qaux.FieldByName("PERMITEAGENCIASEMDV").AsString = "N") Then
       bsShowMessage("É necessário informar dados de conta corrente (Banco,Agencia,Conta-corrente ou DV)", "E")
       CanContinue = False
       Set qaux = Nothing
       Exit Sub
    End If
    Set qaux = Nothing


    If (CurrentQuery.FieldByName("CCNOME").IsNull And (Not CurrentQuery.FieldByName("CCCPFCNPJ").IsNull)) Or _
        ((Not CurrentQuery.FieldByName("CCNOME").IsNull) And CurrentQuery.FieldByName("CCCPFCNPJ").IsNull) Then

      Dim SQLCON As Object
      Dim SQL As Object

      Set SQLCON = NewQuery
      Set SQL = NewQuery


      SQLCON.Active = False
      SQLCON.Add("SELECT BENEFICIARIO, PRESTADOR, PESSOA, TABRESPONSAVEL FROM SFN_CONTAFIN")
      SQLCON.Add("WHERE CONTACORRENTE=:NCONTACORRENTE AND DV=:NDV AND AGENCIA=:NAGENCIA AND BANCO=:NBANCO")
      SQLCON.ParamByName("NCONTACORRENTE").AsString = CurrentQuery.FieldByName("CONTACORRENTE").AsString
      SQLCON.ParamByName("NDV").AsString = CurrentQuery.FieldByName("DV").AsString
      SQLCON.ParamByName("NAGENCIA").AsInteger = CurrentQuery.FieldByName("AGENCIA").AsInteger
      SQLCON.ParamByName("NBANCO").AsInteger = CurrentQuery.FieldByName("BANCO").AsInteger
      SQLCON.Active = True

      If SQLCON.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then 'BENEFICIÁRIO
        SQL.Active = False
        SQL.Add("SELECT B.NOME, B.MATRICULA, M.CPF COD FROM SAM_BENEFICIARIO B, SAM_MATRICULA M")
        SQL.Add("WHERE B.MATRICULA=M.HANDLE AND B.HANDLE=:HBENEFICIARIO")
        SQL.ParamByName("HBENEFICIARIO").AsInteger = SQLCON.FieldByName("BENEFICIARIO").AsInteger
        SQL.Active = True
      ElseIf SQLCON.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then 'PRESTADOR
        SQL.Active = False
        SQL.Add("SELECT CPFCNPJ COD, NOME FROM SAM_PRESTADOR WHERE HANDLE=:HPRESTADOR")
        SQL.ParamByName("HPRESTADOR").AsInteger = SQLCON.FieldByName("PRESTADOR").AsInteger
        SQL.Active = True
      ElseIf SQLCON.FieldByName("TABRESPONSAVEL").AsInteger = 3 Then 'PESSOA
        SQL.Active = False
        SQL.Add("SELECT CNPJCPF COD, NOME FROM SFN_PESSOA WHERE HANDLE=:HPESSOA")
        SQL.ParamByName("HPESSOA").AsInteger = SQLCON.FieldByName("PESSOA").AsInteger
        SQL.Active = True
      End If

      If Not SQL.EOF Then
        If ((CurrentQuery.State = 3) Or (CurrentQuery.State = 2)) Then
          If (CurrentQuery.FieldByName("CCNOME").AsString <> SQL.FieldByName("NOME").AsString) Or (CurrentQuery.FieldByName("CCCPFCNPJ").AsString <> SQL.FieldByName("COD").AsString) Then
            If bsShowMessage("Nome e/ou CPF/CNPJ digitado, diferente do Responsável pela Conta Corrente" + Chr(13) + "Deseja trocar o 'NOME' e/ou 'CPF/CNPJ' do Responsável ?", "Q") = vbYes Then
              CurrentQuery.FieldByName("CCNOME").AsString = SQL.FieldByName("NOME").AsString
              CurrentQuery.FieldByName("CCCPFCNPJ").AsString = SQL.FieldByName("COD").AsString
            End If
          End If
        End If
      End If

      Set SQLCON = Nothing
      Set SQL = Nothing

    End If

    If (CurrentQuery.FieldByName("CCNOME").IsNull And (Not CurrentQuery.FieldByName("CCCPFCNPJ").IsNull)) Or _
        ((Not CurrentQuery.FieldByName("CCNOME").IsNull) And CurrentQuery.FieldByName("CCCPFCNPJ").IsNull) Then
      bsShowMessage("É obrigatório informar nome e cpf/cnpj, caso algum desses campos serem preenchidos.", "E")
      CanContinue = False
      Exit Sub
    End If


    Dim Msg As String

    CNPJCPF = CurrentQuery.FieldByName("CCCPFCNPJ").AsString

    If CNPJCPF <> "" Then
      If Len(CNPJCPF) = 11 Then
        If Not IsValidCPF(CNPJCPF)Then
          bsShowMessage("CPF Inválido.", "E")
          CanContinue = False
          Exit Sub
        End If
      ElseIf Len(CNPJCPF) = 14 Then
        If Not IsValidCGC(CNPJCPF)Then
          bsShowMessage("CNPJ Inválido.", "E")
          CanContinue = False
          Exit Sub
        End If
      Else
        bsShowMessage("CPF/CNPJ Inválido.", "E")
        CanContinue = False
        Exit Sub
      End If

      If Not CheckCPFCNPJ(RetornaNumeros(CurrentQuery.FieldByName("CCCPFCNPJ").AsString), 0, True, Msg) Then
        bsShowMessage(Msg, "E")
        CCCPFCNPJ.SetFocus
        CanContinue = False
        Exit Sub
      End If
    End If

    'verifica se todos os campos da conta corrente foram preenchidos
    Dim vCamposPreenchidos As String
    vCamposPreenchidos = vCamposPreenchidos + Mid(CurrentQuery.FieldByName("BANCO").AsString, 1, 1)
    vCamposPreenchidos = vCamposPreenchidos + Mid(CurrentQuery.FieldByName("AGENCIA").AsString, 1, 1)
    vCamposPreenchidos = vCamposPreenchidos + Mid(CurrentQuery.FieldByName("CONTACORRENTE").AsString, 1, 1)

    Dim qaux2 As Object
    Set qaux2 = NewQuery
    qaux2.Clear
    qaux2.Add("SELECT PERMITEAGENCIASEMDV FROM SFN_BANCO WHERE HANDLE = :HANDLE")
    qaux2.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BANCO").AsInteger
    qaux2.Active = True
    If qaux2.FieldByName("PERMITEAGENCIASEMDV").AsString = "S" Then
      vCamposPreenchidos = vCamposPreenchidos + "1"
    Else
      vCamposPreenchidos = vCamposPreenchidos + Mid(CurrentQuery.FieldByName("DV").AsString, 1, 1)
    End If
    Set qaux2 = Nothing
    If (Len(vCamposPreenchidos) < 4) And (Len(vCamposPreenchidos) > 0) Then
      bsShowMessage("Dados da conta corrente incompletos.", "E")
      CanContinue = False
      Exit Sub
    End If

    If CurrentQuery.FieldByName("CCCPFCNPJ").AsString <> "" Then
      If Not CheckCPFCNPJ(RetornaNumeros(CurrentQuery.FieldByName("CCCPFCNPJ").AsString), 0, True, Msg) Then
        bsShowMessage(Msg, "E")
        CCCPFCNPJ.SetFocus
        CanContinue = False
        Exit Sub
      End If
    End If

    Set interface = CreateBennerObject("FINANCEIRO.CONTAFIN")

    vBanco = CurrentQuery.FieldByName("BANCO").AsInteger
    vAgencia = CurrentQuery.FieldByName("AGENCIA").AsInteger
    vConta = CurrentQuery.FieldByName("CONTACORRENTE").AsString
    vDV = CurrentQuery.FieldByName("DV").AsString

    Dim vsMensagem As String

    vRetorno = interface.VerificaDV(CurrentSystem, vBanco, vAgencia, CurrentQuery.TQuery, vConta, vDV, vsMensagem)
    If Not vRetorno Then
      bsShowMessage(vsMensagem, "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim vPermitidas(31) As String
  Dim I As Integer
  Dim vDV1 As String
  Dim vDV2 As String

  vPermitidas(0) = "+"
  vPermitidas(1) = "-"
  vPermitidas(2) = "*"
  vPermitidas(3) = "/"
  vPermitidas(4) = "."
  vPermitidas(5) = ","
  vPermitidas(6) = ";"
  vPermitidas(7) = ":"
  vPermitidas(8) = "'"
  vPermitidas(9) = "="
  vPermitidas(10) = "|"
  vPermitidas(11) = "_"
  vPermitidas(12) = ")"
  vPermitidas(13) = "("
  vPermitidas(14) = "%"
  vPermitidas(15) = "$"
  vPermitidas(16) = "#"
  vPermitidas(17) = "@"
  vPermitidas(18) = "!"
  vPermitidas(19) = "?"
  vPermitidas(20) = "~"
  vPermitidas(21) = "`"
  vPermitidas(22) = """"
  vPermitidas(23) = "{"
  vPermitidas(24) = "}"
  vPermitidas(25) = "["
  vPermitidas(26) = "]"
  vPermitidas(27) = "^"
  vPermitidas(28) = "\"
  vPermitidas(29) = ">"
  vPermitidas(30) = "<"

  vDV1 = Mid(CurrentQuery.FieldByName("DV").AsString, 1, 1)
  vDV2 = Mid(CurrentQuery.FieldByName("DV").AsString, 2, 1)

  For I = 0 To 30
    If (vDV1 = vPermitidas(I)) Or (vDV2 = vPermitidas(I)) Then
      bsShowMessage("Dígito verificador da conta corrente não pode conter caracteres especiais!", "E")
      CanContinue = False
      Exit Sub
    End If
  Next I

  Set interface = Nothing 'Danilo Raisi SMS: 73239 - 06/12/2006

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Msg As String
  Dim qContaFin As Object
  Set qContaFin = NewQuery

  qContaFin.Clear
  qContaFin.Add("SELECT TABRESPONSAVEL")
  qContaFin.Add("FROM SFN_CONTAFIN")
  qContaFin.Add("WHERE HANDLE = :HCONTAFIN")
  qContaFin.ParamByName("HCONTAFIN").Value = RecordHandleOfTable("SFN_CONTAFIN")
  qContaFin.Active = True
  'Somente se o responsável for um prestador
  If qContaFin.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then
    If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
      Set qContaFin = Nothing
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  Set qContaFin = Nothing

  If CurrentQuery.FieldByName("SITUACAO").Value = "C" Then
    bsShowMessage("Esta alteração já está confirmada.", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  Dim Msg As String
  Dim qContaFin As Object
  Set qContaFin = NewQuery

  qContaFin.Clear
  qContaFin.Add("SELECT TABRESPONSAVEL")
  qContaFin.Add("FROM SFN_CONTAFIN")
  qContaFin.Add("WHERE HANDLE = :HCONTAFIN")
  qContaFin.ParamByName("HCONTAFIN").Value = RecordHandleOfTable("SFN_CONTAFIN")
  qContaFin.Active = True
  'Somente se o responsável for um prestador
  If qContaFin.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then
    If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
      Set qContaFin = Nothing
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  Set qContaFin = Nothing

  If CurrentQuery.FieldByName("SITUACAO").Value = "C" Then
    bsShowMessage("Esta alteração já está confirmada.", "E")
    CanContinue = False
    Exit Sub
  End If

  'André - SMS 24062 - 31/08/2004

  Dim qCF As Object
  Dim qConta As Object

  Set qCF = NewQuery
  Set qConta = NewQuery

  qCF.Clear
  qCF.Add (" SELECT DISTINCT CONTAFINANCEIRA                                     ")
  qCF.Add ("   FROM SFN_FATURA FAT                                               ")
  qCF.Add ("        JOIN SFN_CONTAFIN CON On (FAT.CONTAFINANCEIRA = CON.HANDLE)  ")
  qCF.Add (" WHERE CON.PRESTADOR = :PRESTADOR                                    ")

  qCF.ParamByName("PRESTADOR").Value = RecordHandleOfTable("SAM_PRESTADOR")
  qCF.Active = True

  qConta.Clear
  qConta.Add (" SELECT * FROM SFN_CONTAFIN_ALTERACAO WHERE CONTAFINANCEIRA = :CONTAFIN And SITUACAO = 'S'")

  qConta.ParamByName("CONTAFIN").Value = qCF.FieldByName("CONTAFINANCEIRA").AsInteger
  qConta.Active = True

  viBanco = qConta.FieldByName("BANCO").AsInteger
  viAgencia = qConta.FieldByName("AGENCIA").AsInteger
  VCC = qConta.FieldByName("CONTACORRENTE").AsString
  vDV = qConta.FieldByName("DV").AsString

  'FIM - SMS 24062

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Msg As String
  Dim qContaFin As Object
  Set qContaFin = NewQuery

  qContaFin.Clear
  qContaFin.Add("SELECT TABRESPONSAVEL")
  qContaFin.Add("FROM SFN_CONTAFIN")
  qContaFin.Add("WHERE HANDLE = :HCONTAFIN")
  qContaFin.ParamByName("HCONTAFIN").Value = RecordHandleOfTable("SFN_CONTAFIN")
  qContaFin.Active = True
  'Somente se o responsável for um prestador
  If qContaFin.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then
    If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
      Set qContaFin = Nothing
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  Set qContaFin = Nothing
End Sub

Public Function RetornaNumeros(Strx As String) As String
  Dim I As Long
  RetornaNumeros = ""
  For I = 1 To Len(Strx)
    If InStr("1234567890", Mid(Strx, I, 1)) > 0 Then 'If InStr(".,<>´=-_+/*[]{}?:*()&¨^`%$#@!\|", Mid(Strx, i, 1)) = 0 Then
      RetornaNumeros = RetornaNumeros + Mid(Strx, I, 1)
    End If
  Next I
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCONFIRMA"
			BOTAOCONFIRMA_OnClick
		Case "BOTAONOVOTESOURARIA"
			'NovoTesouraria
	End Select
End Sub

Public Sub AlterarBancoTesouraria(IhandleBanco As Integer)
	CurrentQuery.FieldByName("AGENCIA").Clear
    CurrentQuery.FieldByName("CONTACORRENTE").AsString = ""
    CurrentQuery.FieldByName("DV").AsString = ""
    CurrentQuery.FieldByName("BANCO").AsInteger = IhandleBanco
End Sub
