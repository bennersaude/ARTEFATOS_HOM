'HASH: CCE26B42D309D3C521740208D3053E3D
'Macro: SFN_CONTAFIN
'#Uses "*bsShowMessage"


Option Explicit
Dim NaoGerarDocumentoAnterior As String

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

  vHandle = interface.Exec(CurrentSystem, "SFN_AGENCIA", vColunas, 1, vCampos, vCriterio, "Tabela de Agências", True, "")

  If vHandle <>0 Then
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

  vHandle = interface.Exec(CurrentSystem, "SFN_BANCO", vColunas, 1, vCampos, vCriterio, "Tabela de Bancos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("BANCO").Value = vHandle
  End If
  Set interface = Nothing
End Sub

Public Sub BOTAOADEQUACAO_OnClick()
  Dim interface As Object
  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
    Set interface = CreateBennerObject("CA027.SOLADEQUACAODEB")
    interface.Executar(CurrentSystem, -99, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger)
    Set interface = Nothing
  Else
    bsShowMessage("Adequação de débito somente permitida para Beneficiários !", "I")
  End If
End Sub

Public Sub BOTAOADIANTAMENTO_OnClick()
  Dim OLEAdiantamento As Object
  Set OLEAdiantamento = CreateBennerObject("samadiantamento.rotinas")
  OLEAdiantamento.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set OLEAdiantamento = Nothing
End Sub


Public Sub BOTAOCONSULTA_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SamContaFinanceira.Consulta")
  interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing
End Sub

Public Sub BOTAOCONSULTACAPACIDADEPGTO_OnClick()
  Dim interface As Object
  Dim pCapacidadePagtoFolha As Double
  Dim pSomaTotalPF As Double

  Set interface = CreateBennerObject("Ca043.Autorizacao")
  interface.CalcularSaldoContaFinanceira(CurrentSystem, ServerDate, 0, CurrentQuery.FieldByName("BENEFICIARIO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, pSomaTotalPF, pCapacidadePagtoFolha)
  bsShowMessage("Capacidade de pagamento: " + Format(pCapacidadePagtoFolha, "###,###,##0.00;(###,###,##0.00)") + Chr(13) + _
         "Valor utilizado até o momento: " + Format(pSomaTotalPF, "###,###,##0.00;(###,###,##0.00)") + Chr(13) + _
         "Saldo: " + Format(pCapacidadePagtoFolha-pSomaTotalPF, "###,###,##0.00;(###,###,##0.00)"), "I")
  'MsgBox "Saldo hoje (" + Format(Now) + ") é " + Format(valor, "$ ###,###,##0.00;(###,###,##0.00)")
  Set interface = Nothing
End Sub

Public Sub BOTAOEXCLUI_OnClick()
  Dim vValor As Double
  Dim interface As Object
  Set interface = CreateBennerObject("SAMIMPOSTOS.rotinas")
  interface.CALCULARINSS(CurrentSystem, 41483, CurrentQuery.FieldByName("HANDLE").AsInteger, 2, 16, Date, 500, vValor)

  bsShowMessage(Str(vValor), "I")
  Set interface = Nothing
End Sub

Public Sub BOTAOFATURAAVULSA_OnClick()

  Dim interface As Object
'  Set interface = CreateBennerObject("sfnfatura.rotinas")
'  interface.faturaavulsa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = CreateBennerObject("ESPECIFICO.UESPECIFICO")
  interface.FIN_FaturaAvulsa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set interface = Nothing
End Sub

Public Sub BOTAOPARCELAMENTO_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("parcelamento.rotinas")
  interface.teste(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set interface = Nothing


End Sub

Public Sub BOTAOSALDO_OnClick()

  Dim Intergace As Object
  Dim sALDO As Double
  Dim valor As Currency
  Dim interface As Object

  Set interface = CreateBennerObject("FINANCEIRO.ContaFin")
  sALDO = interface.Saldo(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, Now)
  valor = sALDO
  bsShowMessage( "Saldo hoje (" + Format(Now) + ") é " + Format(valor, "$ ###,###,##0.00;(###,###,##0.00)"), "I")
  Set interface = Nothing

End Sub

Public Sub BOTAOGERADOCUMENTO_OnClick()
	If (CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString = "S") Then
		If bsShowMessage("A conta financeira está marcada para não gerar documentos! Continuar?", "Q") = vbYes Then
			SessionVar("HCONTAFIN") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)
		End If
	End If
End Sub

Public Sub BOTAOREATIVARCOMFATURAMENTO_OnClick()
  Dim interface As Object
  Dim SQLConta As Object
  Set SQLConta = NewQuery

  SQLConta.Active = False
  SQLConta.Add("SELECT TABRESPONSAVEL  ")
  SQLConta.Add("  FROM SFN_CONTAFIN    ")
  SQLConta.Add(" WHERE HANDLE = :HANDLE")
  SQLConta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQLConta.Active = True

  If SQLConta.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then
    Set interface = CreateBennerObject("BSBEN008.Geral")
    interface.Reativar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set interface = Nothing
  Else
    bsShowMessage("Reativação só para beneficiários!", "I")
  End If

End Sub

Public Sub BOTAOTRANSFERESALDO_OnClick()
  Dim interface As Object
  Set interface = CreateBennerObject("SamContaFinanceira.Transferencia")
  interface.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set interface = Nothing

End Sub

Public Sub TABLE_AfterInsert
  If Not VisibleMode Then
    Exit Sub
  End If

  If(Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull)Then
  CurrentQuery.FieldByName("TABRESPONSAVEL").Value = 1
Else
  If(Not CurrentQuery.FieldByName("PRESTADOR").IsNull)Then
  CurrentQuery.FieldByName("TABRESPONSAVEL").Value = 2
Else
  If(Not CurrentQuery.FieldByName("PESSOA").IsNull)Then
  CurrentQuery.FieldByName("TABRESPONSAVEL").Value = 3
End If
End If
End If
End Sub

Public Sub TABLE_AfterScroll()
  Dim SQLPES As Object
  Dim SQLBEN As Object
  Dim SQLPRE As Object
  Set SQLPES = NewQuery
  Set SQLBEN = NewQuery
  Set SQLPRE = NewQuery

  CAPACIDADEPAGTOFOLHA.Visible = False
  BOTAOCONSULTACAPACIDADEPGTO.Visible = False

  ROTULOCODIGOPARABANCO.Text = "Código para o Banco: " + CurrentQuery.FieldByName("HANDLE").AsString

  If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 1 Then 'Caso for Beneficiário
    SQLBEN.Active = False
    SQLBEN.Add("SELECT NOME, BENEFICIARIO FROM SAM_BENEFICIARIO WHERE HANDLE = :HBENEFICIARIO")
    SQLBEN.ParamByName("HBENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
    SQLBEN.Active = True
    If SQLBEN.EOF Then
      ROTULORESPONSAVEL.Text = "Beneficiário não encontrado"
    Else
      ROTULORESPONSAVEL.Text = "BENEFICIÁRIO: " + Format(SQLBEN.FieldByName("BENEFICIARIO").AsString, "000000\.000000\.00") + " - " + SQLBEN.FieldByName("NOME").AsString
    End If

    CAPACIDADEPAGTOFOLHA.Visible = True
    BOTAOCONSULTACAPACIDADEPGTO.Visible = True

  Else
    If CurrentQuery.FieldByName("TABRESPONSAVEL").AsInteger = 2 Then 'Caso for Prestador
      SQLPRE.Active = False
      SQLPRE.Add("SELECT NOME, PRESTADOR FROM SAM_PRESTADOR WHERE HANDLE = :HPRESTADOR")
      SQLPRE.ParamByName("HPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
      SQLPRE.Active = True
      If SQLPRE.EOF Then
        ROTULORESPONSAVEL.Text = "Prestador não encontrado"
      Else
        ROTULORESPONSAVEL.Text = "PRESTADOR: "
        If Len(SQLPRE.FieldByName("PRESTADOR").AsString)<= 11 Then
          ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + _
                                   Format(SQLPRE.FieldByName("PRESTADOR").AsString, "000\.000\.000\-00")
        Else
          ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + _
                                   Format(SQLPRE.FieldByName("PRESTADOR").AsString, "00\.000\.000\/0000\-00")
        End If
        ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + " - " + _
                                 SQLPRE.FieldByName("NOME").AsString
      End If
    Else 'Caso for Pessoa
      SQLPES.Active = False
      SQLPES.Add("SELECT NOME, CNPJCPF FROM SFN_PESSOA WHERE HANDLE = :HPESSOA")
      SQLPES.ParamByName("HPESSOA").AsInteger = CurrentQuery.FieldByName("PESSOA").AsInteger
      SQLPES.Active = True
      If SQLPES.EOF Then
        ROTULORESPONSAVEL.Text = "Pessoa não encontrada"
      Else
        ROTULORESPONSAVEL.Text = "PESSOA: "
        If Len(SQLPES.FieldByName("CNPJCPF").AsString)<= 11 Then
          ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + _
                                   Format(SQLPES.FieldByName("CNPJCPF").AsString, "000\.000\.000\-00")
        Else
          ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + _
                                   Format(SQLPES.FieldByName("CNPJCPF").AsString, "00\.000\.000\/0000\-00")
        End If
        ROTULORESPONSAVEL.Text = ROTULORESPONSAVEL.Text + " - " + _
                                 SQLPES.FieldByName("NOME").AsString
      End If
    End If
  End If

  SQLBEN.Active = False
  SQLPES.Active = False
  SQLPRE.Active = False
  Set SQLPES = Nothing
  Set SQLPRE = Nothing
  Set SQLBEN = Nothing
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If Not VisibleMode Then
    Exit Sub
  End If

  If RecordHandleOfTable("SAM_PRESTADOR")>0 Then
    Dim Msg As String
    If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
      bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  NaoGerarDocumentoAnterior = CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  NaoGerarDocumentoAnterior = "N"
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Texto As String

  'SMS 49152 - Anderson Lonardoni
  'Esta verificação foi tirada do BeforeInsert e colocada no
  'BeforePost para que, no caso de Inserção, já existam valores
  'no CurrentQuery e para funcionar com o Integrator
  If CurrentQuery.FieldByName("PRESTADOR").AsInteger > 0 Then
    Dim Msg As String
    If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
  	  bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
    'SMS 49152 - Fim
  End If

  If Not VisibleMode Then
    If CurrentQuery.State = 2 Then
      CurrentQuery.FieldByName("ALTERACAOUSUARIO").Value = CurrentUser
      CurrentQuery.FieldByName("ALTERACAODATA").Value = ServerNow
    Else
      If CurrentQuery.State = 3 Then
        CurrentQuery.FieldByName("INCLUSAOUSUARIO").Value = CurrentUser
        CurrentQuery.FieldByName("INCLUSAODATA").Value = ServerNow
      End If
    End If
    Exit Sub
  End If

  If NaoGerarDocumentoAnterior <>CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString Then
    If CurrentQuery.FieldByName("NAOGERARDOCUMENTO").AsString = "S" Then
      Texto = "A conta financeira foi modificada para NÃO GERAR documento ! "
    Else
      Texto = "A conta financeira foi modificada para GERAR documento ! "
    End If
    Texto = Texto + " Confirma ?"
    If bsShowMessage(Texto, "Q")<>vbYes Then
      CanContinue = False
    End If
  End If

  '=================================================================================
  'Cálculo do dígito da Conta Corrente -** Junior 06/04/2000
  '=================================================================================
  If CurrentQuery.FieldByName("TABGERACAOPAG").AsInteger = 1 Or CurrentQuery.FieldByName("TABGERACAOREC").AsInteger = 1 Then

    'sms 33910
    If CurrentQuery.FieldByName("BANCO").IsNull Or _
       CurrentQuery.FieldByName("AGENCIA").IsNull Or _
       CurrentQuery.FieldByName("CONTACORRENTE").IsNull Or _
       CurrentQuery.FieldByName("DV").IsNull Then
       '>>>> Comentado por Luciano T. Alberti - SMS 93931 - 28/02/2008
       'CurrentQuery.FieldByName("CCNOME").IsNull Or _
       'CurrentQuery.FieldByName("CCCPFCNPJ").IsNull Then
       '<<<<<<
      bsShowMessage("É necessário informar dados da conta corrente", "E")
      CanContinue = False
      Exit Sub
    End If


    Dim SQL As Object

    Set SQL = NewQuery

    SQL.Add("SELECT CODIGO")
    SQL.Add("FROM SFN_BANCO")
    SQL.Add("WHERE HANDLE = :HBANCO")
    SQL.ParamByName("HBANCO").Value = CurrentQuery.FieldByName("BANCO").AsInteger
    SQL.Active = True

    If SQL.EOF Then
      Set SQL = Nothing
      bsShowMessage("Banco não encontrado", "E")
      CanContinue = False
    Else
      Select Case SQL.FieldByName("CODIGO").AsInteger
        '=================================================================================
        '************************* BANCO DO NORDESTE  ************************************
        '=================================================================================
        Case 4
          SQL.Clear
          SQL.Add("SELECT DVCONTACORRENTE")
          SQL.Add("FROM SFN_AGENCIA")
          SQL.Add("WHERE HANDLE = :HAGENCIA")
          SQL.ParamByName("HAGENCIA").Value = CurrentQuery.FieldByName("AGENCIA").AsInteger
          SQL.Active = True

          If SQL.EOF Then
            Set SQL = Nothing
            bsShowMessage("Agência não encontrada", "E")
            CanContinue = False
          Else
            If SQL.FieldByName("DVCONTACORRENTE").IsNull Then
              Set SQL = Nothing
              bsShowMessage("Falta dígito para Conta Corrente na Agência", "E")
              CanContinue = False
            Else
              Dim vDigito As Integer
              Dim vCodigoAgencia As String
              Dim vContaCorrente As String
              Dim j As Integer
              Dim I As Integer

              vDigito = 0
              vCodigoAgencia = Format(SQL.FieldByName("DVCONTACORRENTE").AsInteger, "000")
              vContaCorrente = Format(CurrentQuery.FieldByName("CONTACORRENTE").AsInteger, "00000")

              j = 9
              For I = 1 To 3
                vDigito = vDigito + Val(Mid(vCodigoAgencia, I, 1)) * j
                j = j -1
              Next I

              For I = 1 To 5
                vDigito = vDigito + Val(Mid(vContaCorrente, I, 1)) * j
                j = j -1
              Next I

              vDigito = vDigito Mod 11
              If vDigito = 0 Or vDigito = 1 Then
                vDigito = 0
              Else
                vDigito = 11 - vDigito
              End If

              Set SQL = Nothing

              If vDigito <>CurrentQuery.FieldByName("DV").AsInteger Then
                bsShowMessage("Digito da Conta Corrente não confere", "E")
                CanContinue = False
              End If
            End If
          End If

          '=================================================================================Juliano 12/01/01
          '************************* BANCO DO NORDESTE  ************************************
          '=================================================================================

        Case 1
          Dim Resto As Long
          Dim Mult As Long
          Dim Total As Long
          Dim Ind As Long
          Dim Numero As String
          Dim Campo1 As String
          Dim DVConta As String
          Dim Digito As String
          Dim Resultado As Boolean

          Resultado = False
          Numero = CurrentQuery.FieldByName("CONTACORRENTE").AsString
          DVConta = CurrentQuery.FieldByName("DV").AsString
          Campo1 = Right("00000000000000000" + Numero, 17)
          Mult = 2
          Total = 0

          If Trim(Campo1) = "0" Then
            Resultado = False
            bsShowMessage("Conta Corrente está em branco!", "E")
            CanContinue = False
          End If

          For Ind = 17 To 1 Step -1
            Total = Total + (CLng(Mid(Campo1, Ind, 1)) * Mult)
            Mult = Mult + 1
            If Mult = 10 Then
              Mult = 2
            End If
          Next

          Resto = Total Mod 11

          If Resto = 1 Then
            Digito = "X"
          Else
            If Resto = 0 Then
              Digito = "0"
            Else
              Digito = CStr(11 - Resto)
            End If
          End If

          If(DVConta <>Digito)And(DVConta <>"-1")Then
          Resultado = False
        Else
          Resultado = True
        End If

        If Resultado = False Then
          bsShowMessage("Digito da Conta Corrente não confere", "E")
          CanContinue = False
        End If

    End Select

    Set SQL = Nothing
  End If
End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If Not VisibleMode Then
    Exit Sub
  End If
  If RecordHandleOfTable("SAM_PRESTADOR")>0 Then
    Dim Msg As String
    If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
	  bsShowMessage(Msg, "E")
      CanContinue = False
      Exit Sub
    End If
  End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------

Public Function checkPermissaoFilial(CurrentSystem, pServico As String, pTabela As String, pMsg As String)As String

  Dim vFiltro, vResultado As String
  Dim qAuxiliar, qPermissoes, SamPrestadorParametro As Object
  Dim qFilialProc As Object

  Dim SamPrestador

  Set qPermissoes = NewQuery
  Set qFilialProc = NewQuery
  Set qAuxiliar = NewQuery
  Set SamPrestadorParametro = NewQuery

  SamPrestadorParametro.Add(" SELECT CONTROLEDEACESSO, BLOQUEIOFILIALPROCESSAMENTO ")
  SamPrestadorParametro.Add(" FROM SAM_PARAMETROSPRESTADOR  ")
  SamPrestadorParametro.Active = True

  If SamPrestadorParametro.FieldByName("CONTROLEDEACESSO").AsString = "N" Then
    checkPermissaoFilial = "(SELECT HANDLE FROM MUNICIPIOS)"
    pMsg = ""
    Exit Function
  End If

  ' começa o controle de acesso
  ' verifica bloqueio filial de processamento

  If SamPrestadorParametro.FieldByName("BLOQUEIOFILIALPROCESSAMENTO").AsString = "S" Then
    qAuxiliar.Clear
    qAuxiliar.Add("SELECT FILIALPADRAO")
    qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
    qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
    qAuxiliar.Active = True
    qFilialProc.Active = False
    qFilialProc.Add("Select FILIALPROCESSAMENTO FROM FILIAIS WHERE HANDLE = :HANDLE")
    qFilialProc.ParamByName("HANDLE").Value = qAuxiliar.FieldByName("FILIALPADRAO").Value
    qFilialProc.Active = True
    If qFilialProc.FieldByName("FILIALPROCESSAMENTO").AsInteger = qAuxiliar.FieldByName("FILIALPADRAO").Value Then
      checkPermissaoFilial = "N"
      pMsg = "Permissão negada! Filial padrão do usuário igual a sua filial de processamento."
      Exit Function
    End If
  End If

  qAuxiliar.Active = False
  qAuxiliar.Clear
  If pTabela = "P" Then
    qAuxiliar.Add("SELECT FILIALPADRAO")
    qAuxiliar.Add("  FROM SAM_PRESTADOR")
    qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
    qAuxiliar.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
    qAuxiliar.Active = True
    If qAuxiliar.FieldByName("FILIALPADRAO").IsNull Then
      qAuxiliar.Clear
      qAuxiliar.Add("SELECT FILIALPADRAO")
      qAuxiliar.Add("  FROM Z_GRUPOUSUARIOS")
      qAuxiliar.Add(" WHERE HANDLE = :HANDLE")
      qAuxiliar.ParamByName("HANDLE").Value = CurrentUser
      qAuxiliar.Active = True
    End If

    qPermissoes.Active = False
    qPermissoes.Clear
    qPermissoes.Add("SELECT X.ALTERAR, ")
    qPermissoes.Add("       X.INCLUIR, ")
    qPermissoes.Add("       X.EXCLUIR, ")
    qPermissoes.Add("       X.FILIAL ")
    qPermissoes.Add("  FROM (SELECT A.ALTERAR ALTERAR, ")
    qPermissoes.Add("               A.INCLUIR INCLUIR, ")
    qPermissoes.Add("               A.EXCLUIR EXCLUIR, ")
    qPermissoes.Add("               A.FILIAL  FILIAL ")
    qPermissoes.Add("          FROM Z_GRUPOUSUARIOS_FILIAIS A ")
    qPermissoes.Add("         WHERE  A.USUARIO = :USUARIO ")
    qPermissoes.Add("           AND  A.FILIAL  = :FILIAL ")
    qPermissoes.Add("        UNION ")
    qPermissoes.Add("        SELECT U.ALTERAR      ALTERAR, ")
    qPermissoes.Add("               U.INCLUIR      INCLUIR, ")
    qPermissoes.Add("               U.EXCLUIR      EXCLUIR, ")
    qPermissoes.Add("               U.FILIALPADRAO FILIAL ")
    qPermissoes.Add("          FROM Z_GRUPOUSUARIOS U ")
    qPermissoes.Add("         WHERE U.HANDLE = :USUARIO ")
    qPermissoes.Add("           AND U.FILIALPADRAO  = :FILIAL) X ")

    qPermissoes.ParamByName("USUARIO").Value = CurrentUser
    qPermissoes.ParamByName("FILIAL").Value = qAuxiliar.FieldByName("FILIALPADRAO").AsInteger
    qPermissoes.Active = True
  End If


  If pServico = "A" Then
    ' Verifica se pode alterar conforme a filial padrao
    vFiltro = vFiltro + _
              "(SELECT DISTINCT M.HANDLE " + _
              "  FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "       SAM_REGIAO R, " + _
              "       MUNICIPIOS M " + _
              " WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "   AND R.FILIAL = A.FILIAL " + _
              "   AND M.REGIAO = R.HANDLE " + _
              "   AND A.ALTERAR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.ALTERAR = 'S' ) "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode alterar
    vFiltro = ""
    vFiltro = vFiltro + _
              "(SELECT DISTINCT M.HANDLE " + _
              "   FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "        MUNICIPIOS M, " + _
              "        SAM_REGIAO R " + _
              "  WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "    AND M.REGIAO = R.HANDLE " + _
              "    AND A.FILIAL = R.FILIAL " + _
              "    AND A.ALTERAR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.ALTERAR = 'S' )"
  End If
  If pServico = "I" Then
    ' Verifica se pode incluir conforme a filial padrao
    vFiltro = vFiltro + _
              "(SELECT DISTINCT M.HANDLE " + _
              "  FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "       SAM_REGIAO R, " + _
              "       MUNICIPIOS M " + _
              " WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "   AND R.FILIAL = A.FILIAL " + _
              "   AND M.REGIAO = R.HANDLE " + _
              "   AND A.INCLUIR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE" + _
              "    FROM Z_GRUPOUSUARIOS U," + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.INCLUIR = 'S' ) "

    qAuxiliar.Active = False
    qAuxiliar.Clear
    qAuxiliar.Add(vFiltro)
    qAuxiliar.Active = True
    ' Retorna o filtro dos municipios que pode incluir
    vFiltro = ""
    vFiltro = vFiltro + _
              "(Select DISTINCT M.HANDLE " + _
              "   FROM Z_GRUPOUSUARIOS_FILIAIS A, " + _
              "        MUNICIPIOS M, " + _
              "        SAM_REGIAO R " + _
              "  WHERE A.USUARIO = " + CStr(CurrentUser) + _
              "    AND M.REGIAO = R.HANDLE " + _
              "    AND A.FILIAL = R.FILIAL " + _
              "    AND A.INCLUIR = 'S' " + _
              " UNION " + _
              "  SELECT M.HANDLE " + _
              "    FROM Z_GRUPOUSUARIOS U, " + _
              "         SAM_REGIAO R,  " + _
              "         MUNICIPIOS M  " + _
              "   WHERE U.HANDLE = " + CStr(CurrentUser) + _
              "     AND R.FILIAL = U.FILIALPADRAO " + _
              "     AND M.REGIAO = R.HANDLE " + _
              "     AND U.INCLUIR = 'S' ) "

  End If

  ' se não estiver cadastrado
  If(qPermissoes.FieldByName("ALTERAR").IsNull)Then
  If pServico = "" Then
    checkPermissaoFilial = ""
    Exit Function
  End If
End If

' se não informou o servico,retorna uma String com os servicos permitidos "LAIE"
If(pServico = "")Then
vResultado = ""
If qPermissoes.FieldByName("ALTERAR").AsString = "S" Then
  vResultado = vResultado + "A"
End If
If qPermissoes.FieldByName("INCLUIR").AsString = "S" Then
  vResultado = vResultado + "I"
End If
If qPermissoes.FieldByName("EXCLUIR").AsString = "S" Then
  vResultado = vResultado + "E"
End If
' se informou o servico,retorna S/N
Else
  Select Case pServico
    Case "A"
      If qPermissoes.FieldByName("ALTERAR").AsString = "S" Then
        vResultado = "S"
        If(Not qAuxiliar.FieldByName("Handle").IsNull)Then
        vResultado = vFiltro
      Else
        vResultado = "N"
        pMsg = "Permissão negada! Usuário não pode alterar."
      End If
    Else
      vResultado = "N"
      pMsg = "Permissão negada! Usuário não pode alterar."
    End If
  Case "I"
    If qPermissoes.FieldByName("INCLUIR").AsString = "S" Then
      vResultado = "S"
      If(Not qAuxiliar.FieldByName("Handle").IsNull)Then
      vResultado = vFiltro
    Else
      vResultado = "N"
      pMsg = "ermissão negada! Usuário não pode incluir."
    End If
  Else
    vResultado = "N"
    pMsg = "Permissão negada! Usuário não pode incluir."
  End If
Case "E"
  If qPermissoes.FieldByName("EXCLUIR").AsString = "S" Then
    vResultado = "S"
  Else
    vResultado = "N"
    pMsg = "Permissão negada! Usuário não pode excluir."
  End If
End Select
End If
checkPermissaoFilial = vResultado
End Function


Public Function BuscarFiliais(CurrentSystem, prFilial As Long, prFilialProcessamento As Long, prMsg As String)As Boolean

  Dim qPermissoes As Object
  Set qPermissoes = NewQuery

  BuscarFiliais = True
  qPermissoes.Active = False
  qPermissoes.Clear
  qPermissoes.Add("SELECT A.HANDLE, A.FILIALPROCESSAMENTO")
  qPermissoes.Add("FROM   Z_GRUPOUSUARIOS U,             ")
  qPermissoes.Add("       FILIAIS A                      ")
  qPermissoes.Add("WHERE  (U.HANDLE = :USUARIO)          ")
  qPermissoes.Add("AND    (A.HANDLE = U.FILIALPADRAO)    ")
  qPermissoes.ParamByName("USUARIO").Value = CurrentUser
  qPermissoes.Active = True

  If qPermissoes.EOF Then
    prMsg = "Problemas Usuario x Filial."
    Exit Function
  End If

  prFilial = qPermissoes.FieldByName("HANDLE").AsInteger
  prFilialProcessamento = qPermissoes.FieldByName("FILIALPROCESSAMENTO").AsInteger
  prMsg = ""
  BuscarFiliais = False
  Set qPermissoes = Nothing
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOADEQUACAO"
			BOTAOADEQUACAO_OnClick
		Case "BOTAOADIANTAMENTO"
			BOTAOADIANTAMENTO_OnClick
		Case "BOTAOCONSULTA"
			BOTAOCONSULTA_OnClick
		Case "BOTAOCONSULTACAPACIDADEPGTO"
			BOTAOCONSULTACAPACIDADEPGTO_OnClick
		Case "BOTAOEXCLUI"
			BOTAOEXCLUI_OnClick
		Case "BOTAOFATURAAVULSA"
			BOTAOFATURAAVULSA_OnClick
		Case "BOTAOPARCELAMENTO"
			BOTAOPARCELAMENTO_OnClick
		Case "BOTAOSALDO"
			BOTAOSALDO_OnClick
		Case "BOTAOREATIVARCOMFATURAMENTO"
			BOTAOREATIVARCOMFATURAMENTO_OnClick
		Case "BOTAOTRANSFERESALDO"
			BOTAOTRANSFERESALDO_OnClick
	End Select
End Sub
