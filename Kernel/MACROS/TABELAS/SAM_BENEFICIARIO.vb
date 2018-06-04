'HASH: 5B37C6D49940870FEE23A03DE8D42458
'MACRO: SAM_BENEFICIARIO
'#Uses "*bsShowMessage"
'#Uses "*RegistrarLogAlteracao"

Option Explicit

Dim gbGerarCartao As Boolean
Dim gbPermiteDependenteTitular As Boolean
Dim gbAchouCodigoTabelaPrcTitular As Boolean
Dim gbIsentouCarenciaRecemNascido As Boolean
Dim gbDescontouCarenciasAdotivo As Boolean
Dim gbCriticaEmissaoCartao As Boolean
Dim gbAlteracaoCadastral As Boolean
Dim gdDataFechamento As Date
Dim gdDataNascimento As Date
Dim gdDataAdocao As Date
Dim gdDataAdmissaoTitular As Date
Dim gdAdesaoTitular As Date
Dim gdDataAdesaoDireitoPlano As Date
Dim gdDataCasamento As Date
Dim gsCodigoTabelaPrcTitular As String
Dim gsNaoFaturarModulosAnterior As String
Dim gsNaoFaturarGuiasAnterior As String
Dim gsVerificarContribPrevidencia As String
Dim gsTitularAutonomo As String
Dim gsNaoTemCarenciaAnterior As String
Dim gsGrupoDependenteAnterior As String
Dim gsCodigoTabelaPrcTitularAnterior As String
Dim gsEnderecoCorrespondencia As String
Dim gsMatriculaDigitada As String
Dim gsControleCarencia As String
Dim gsSexo As String
Dim giHTipoDependenteAnterior As Long
Dim giHEstadoCivilAnterior As Long
Dim giPrazoAdesaoRecemNascidosAdotivos As Long
Dim giMotivoBloqueioBenef As Long
Dim giDependente As Long
Dim giOrigemCarenciaTitular As Long
Dim giDiasCompraCarenciaTitular As Long
Dim giTitularConjuge As Long
Dim giUsuarioLiberou As Long
Dim giDiasCompraCarenciaAnterior As Long
Dim giMatriculaIndicadora As Long
Dim giIndiceBusca As Long
Dim giHandleMatricula As Long

Dim vsModoEdicao As String

Dim giUltimoBeneficiarioEndereco As Long
Dim vsXMLContainerEnderecos As String
Dim vsXMLEnderecosExcluidos As String

Dim gsSituacaoRhAux As String
Dim gsDataSituacaoRh As String

Public Sub BENEFICIARIOINDICADOR_OnChange()
  If CurrentQuery.FieldByName("BENEFICIARIOINDICADOR").IsNull Then
      MATRICULAINDICADORA.ReadOnly = False
    End If
End Sub

Public Sub BENEFICIARIOINDICADOR_OnPopup(ShowPopup As Boolean)
  Dim vsTitulo As String
  Dim vsGrid As String
  Dim vsColunas As String
  Dim vsTabela As String
  Dim vsWhere As String
  Dim viHandle As Long
  Dim dllProcura As Object
  Dim qParametrosBeneficiario As BPesquisa

  Set qParametrosBeneficiario = NewQuery
  qParametrosBeneficiario.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")
  qParametrosBeneficiario.Active = True

  ShowPopup = False

  vsTitulo = "Procura por Beneficiário"
  vsGrid = "Nome|Beneficiario|Titular|Matrícula|CPF|RG"
  vsColunas = "SAM_BENEFICIARIO.Z_NOME|SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.EHTITULAR|M.MATRICULA|M.CPF|M.RG"
  vsTabela = "SAM_BENEFICIARIO|SAM_MATRICULA M[SAM_BENEFICIARIO.MATRICULA = M.HANDLE]"

  vsWhere = "EXISTS(SELECT *" _
      + "         FROM SAM_CONTRATO C" _
      + "        WHERE SAM_BENEFICIARIO.CONTRATO = C.HANDLE" _
      + "          AND C.PERMITEINDICAR = 'S')" _
      + "AND SAM_BENEFICIARIO.EHTITULAR = 'S'"

  If (qParametrosBeneficiario.FieldByName("PERMITEINDICADORCANCELADO").AsString = "N") Then vsWhere = vsWhere + " AND SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL"

    Set dllProcura = CreateBennerObject("PROCURA.Procurar")

  viHandle = dllProcura.Exec(CurrentSystem, _
                 vsTabela, _
                 vsColunas, _
                 1, _
                 vsGrid, _
                 vsWhere, _
                 vsTitulo, _
                 False, _
                 "")

  Set dllProcura = Nothing

  If (viHandle <> 0) Then
    If (CurrentQuery.State = 2) Or _
       (CurrentQuery.State = 3) Then
      CurrentQuery.FieldByName("BENEFICIARIOINDICADOR").AsInteger = viHandle
    Else
      bsShowMessage("O registro não está em edição!", "I")
    End If

    Dim qAux As BPesquisa

    Set qAux = NewQuery

    qAux.Active = False

    qAux.Clear

    qAux.Add("SELECT MATRICULA     ")
    qAux.Add("  FROM SAM_BENEFICIARIO")
    qAux.Add(" WHERE HANDLE = :HANDLE")

    qAux.ParamByName("HANDLE").AsInteger = viHandle

    qAux.Active = True

    CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger = qAux.FieldByName("MATRICULA").AsInteger

    If (CurrentQuery.FieldByName("EHTITULAR").AsString = "S") Then giMatriculaIndicadora = qAux.FieldByName("MATRICULA").AsInteger

    MATRICULAINDICADORA.ReadOnly = True
  End If
End Sub

Public Sub BOTAOALTERARADESAO_OnClick()
  Dim vcContainer     As Object
  Dim BSINTERFACE0002 As Object
  Dim vsMensagem      As String

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  Dim viRetorno As Long
  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
                   1, _
                   "TV_FORM0086", _
                   "Alterar Data Adesão", _
                   0, _
                   300, _
                   310, _
                   False, _
                   vsMensagem, _
                   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
  Case -1
    bsShowMessage("Operação cancelada pelo usuário!", "I")
  Case 1
    bsShowMessage(vsMensagem, "I")
  End Select
  Set BSINTERFACE0002 = Nothing
End Sub

Public Sub BOTAOANOTADM_OnClick()
  Dim dllBSInterface0052 As Object
  Set dllBSInterface0052 = CreateBennerObject("BSInterface0052.AnotAdmBenef")
  dllBSInterface0052.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set dllBSInterface0052 = Nothing
End Sub

Public Sub BOTAOBENEFICIARIO_OnClick()
  If (CurrentQuery.FieldByName("HANDLE").AsInteger = 0) Then
    bsShowMessage("É necessário selecionar um beneficiário", "I")

    Exit Sub
  End If

  If (CurrentQuery.State = 2) Or _
     (CurrentQuery.State = 3) Then
    bsShowMessage("O registro não pode estar em edição!", "I")

    Exit Sub
  End If

  Dim dllSamConsultaBenef As Object

  Set dllSamConsultaBenef = CreateBennerObject("SAMCONSULTABENEF.Consultas")

  dllSamConsultaBenef.Executar(CurrentSystem, _
                     3, _
                   0, _
                   0, _
                   CurrentQuery.FieldByName("HANDLE").AsInteger)

    Set dllSamConsultaBenef = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
                   1, _
                   "TV_FORM0019", _
                   "Cancelamento de beneficiário", _
                   0, _
                   300, _
                   420, _
                   False, _
                   vsMensagem, _
                   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
  Case -1
    bsShowMessage("Operação cancelada pelo usuário!", "I")
  Case 1
      bsShowMessage(vsMensagem, "I")
  End Select

  Set BSINTERFACE0002 = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOCARTAO_OnClick()
  Dim vcContainer     As Object
  Dim BSINTERFACE0002 As Object
  Dim vsMensagem      As String

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  Dim viRetorno As Long
  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
                   1, _
                   "TV_FORM0013", _
                   "Gerar Cartão", _
                   0, _
                   140, _
                   310, _
                   False, _
                   vsMensagem, _
                   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
  Case -1
    bsShowMessage("Operação cancelada pelo usuário!", "I")
  Case 1
    bsShowMessage(vsMensagem, "I")
  End Select
  Set BSINTERFACE0002 = Nothing
End Sub

Public Sub BOTAOENDERECO_OnClick()
  Dim dllBSInterface0028           As Object
    Dim vsMensagem                   As String
    Dim viHEnderecoResidencial       As Long
    Dim viHEnderecoComercial         As Long
    Dim viHEnderecoCorrespondencia   As Long
    Dim viHEnderecoAtendimentoDomi   As Long
    Dim pErrNum           As Long
    Dim pErrDesc           As String
    Dim vIniciouTransacao       As Boolean
    Dim vsXMLEnderecoAuxiliar       As String
    Dim vsXMLExcluidosAuxiliar      As String
    Dim vbFecharTransacao           As Boolean
  vIniciouTransacao = False

    viHEnderecoResidencial     = CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger
    viHEnderecoComercial       = CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger
    viHEnderecoCorrespondencia = CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger
    viHEnderecoAtendimentoDomi = CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger

    Set dllBSInterface0028 = CreateBennerObject("BSInterface0028.Endereco")

    'Se for inclusão de registro deve passar zero para o parâmetro de beneficiário
    'para que não se tente fazer buscas ou tratamento sobre um registro que ainda não existe

   If CurrentQuery.State = 3 Then
       giUltimoBeneficiarioEndereco = 0
    vsXMLContainerEnderecos = ""
    vsXMLEnderecosExcluidos = ""
     Else
       If (giUltimoBeneficiarioEndereco > 0) And (giUltimoBeneficiarioEndereco <> CurrentQuery.FieldByName("HANDLE").AsInteger) Then
      vsXMLContainerEnderecos = ""
      vsXMLEnderecosExcluidos = ""
    Else
      If vsXMLContainerEnderecos = "Vazio" Then
        vsXMLContainerEnderecos = ""
      End If
      If vsXMLEnderecosExcluidos = "Vazio" Then
        vsXMLEnderecosExcluidos = ""
      End If
    End If
       giUltimoBeneficiarioEndereco = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If

  vbFecharTransacao = True

  If Not InTransaction Then
    StartTransaction
    vIniciouTransacao = True
  End If
  On Error GoTo Except
    Dim vResultado As Long
    vResultado = dllBSInterface0028.Beneficiario(CurrentSystem, _
                                       giUltimoBeneficiarioEndereco, _
                                       viHEnderecoResidencial, _
                                       viHEnderecoComercial, _
                                       viHEnderecoCorrespondencia, _
                                       viHEnderecoAtendimentoDomi, _
                                       vsXMLContainerEnderecos, _
                                       vsXMLEnderecosExcluidos, _
                                       vsMensagem)

    Dim vEspecificoDLL As Object
      Set vEspecificoDLL = CreateBennerObject("Especifico.uEspecifico")
    vEspecificoDLL.ATE_ConfirmaAlteracaoBeneficiario(CurrentSystem)
    Set vEspecificoDLL = Nothing

    If vResultado = 1 Then
      vsXMLContainerEnderecos = ""
      vsXMLEnderecosExcluidos = ""
      Err.Raise(1, Err, vsMensagem)
    Else
      Select Case CurrentQuery.State
        Case 1
          'Pessoa não está em edição
          If   (viHEnderecoResidencial <> CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger) Or _
            (viHEnderecoComercial <> CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger) Or _
            (viHEnderecoCorrespondencia  <> CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger) Or _
            (viHEnderecoAtendimentoDomi  <> CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger) Then
                        ' os xmls são apagados ao editar então faz uma cópia para não peder os dados ao editar
                        ' novamente depois de já ter alterado
                        Dim vsXMLContainerEnderecosAuxiliar As String
                        vsXMLContainerEnderecosAuxiliar = vsXMLContainerEnderecos
                        vsXMLExcluidosAuxiliar          = vsXMLEnderecosExcluidos
                        vbFecharTransacao               = False
            CurrentQuery.Edit
            'devolve os valores dos xmls a variáveis que são salvas.
            vsXMLContainerEnderecos = vsXMLContainerEnderecosAuxiliar
            vsXMLEnderecosExcluidos = vsXMLExcluidosAuxiliar
              If viHEnderecoResidencial = 0 Then
                CurrentQuery.FieldByName("ENDERECORESIDENCIAL").Clear
              Else
                CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger = viHEnderecoResidencial
              End If

                If viHEnderecoComercial = 0 Then
                  CurrentQuery.FieldByName("ENDERECOCOMERCIAL").Clear
                Else
                  CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger = viHEnderecoComercial
                End If

                If viHEnderecoCorrespondencia = 0 Then
                  CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").Clear
                Else
                  CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger = viHEnderecoCorrespondencia
                End If

                If viHEnderecoAtendimentoDomi = 0 Then
                  CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").Clear
                Else
                  CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger = viHEnderecoAtendimentoDomi
                End If

          End If
        Case 2, 3
          'Efetua preenchimento com os novos valores
            If viHEnderecoResidencial = 0 Then
              CurrentQuery.FieldByName("ENDERECORESIDENCIAL").Clear
            Else
              CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger = viHEnderecoResidencial
            End If

	        If viHEnderecoComercial = 0 Then
	          CurrentQuery.FieldByName("ENDERECOCOMERCIAL").Clear
	        Else
	          CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger = viHEnderecoComercial
	        End If

            If viHEnderecoCorrespondencia = 0 Then
              CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").Clear
            Else
              CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger = viHEnderecoCorrespondencia
            End If

            If viHEnderecoAtendimentoDomi = 0 Then
              CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").Clear
            Else
              CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger = viHEnderecoAtendimentoDomi
            End If
      End Select
    End If

        If ( CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger > 0) Then
          Dim DLLAtualizaRegiao As Object
          If (CurrentQuery.State = 1) Then
          Set DLLAtualizaRegiao = CreateBennerObject("SAMBENEFICIARIO.Atualiza")
            DLLAtualizaRegiao.Beneficiario(CurrentSystem, _
                                         CurrentQuery.FieldByName("HANDLE").AsInteger,  _
                                         CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger, _
                                         vsMensagem)
          Else
            Dim vSqlMunicipio As BPesquisa
            Set vSqlMunicipio = NewQuery
            vSqlMunicipio.Add("SELECT MUNICIPIO FROM SAM_ENDERECO ")
            vSqlMunicipio.Add(" WHERE HANDLE = :ENDERECO ")
            vSqlMunicipio.ParamByName("ENDERECO").AsInteger = CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger
            vSqlMunicipio.Active = True

            Dim vRegiao As Long
          Set DLLAtualizaRegiao = CreateBennerObject("SAMBENEFICIARIO.Cadastro")
        vRegiao = DLLAtualizaRegiao.Regiao(CurrentSystem, _
                                         vSqlMunicipio.FieldByName("MUNICIPIO").AsInteger,  _
                                         CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger, _
                                         vsMensagem)
           If (vRegiao > 0) And (CurrentQuery.FieldByName("REGIAO").AsInteger <> vRegiao) Then
             CurrentQuery.FieldByName("REGIAO").AsInteger = vRegiao
           End If

           vSqlMunicipio.Active = False
           Set vSqlMunicipio = Nothing
          End If
          Set DLLAtualizaRegiao = Nothing
        End If

    If vIniciouTransacao And vbFecharTransacao And InTransaction Then
      Commit
      vIniciouTransacao = False
    End If

      CarregaRotulosEndereco

    If ( CurrentQuery.State = 1) Then
          CurrentQuery.Active = False
          CurrentQuery.Active = True
        End If

      Exit Sub
  Except:
    pErrNum = Err.Number
    pErrDesc = Err.Description
     Set dllBSInterface0028 = Nothing
      If vIniciouTransacao And vbFecharTransacao And InTransaction Then
        Rollback
        vIniciouTransacao = False
        bsShowMessage("("+ CStr( pErrNum) +")"+pErrDesc, "E") 'vai mostrar a mensagem somente se for erro(estava mostrando quando abortava a dll)
      End If
      CarregaRotulosEndereco
End Sub

Public Sub BOTAOFINANCEIRO_OnClick()
  Dim viContaFinanceira As Long

  viContaFinanceira = RetornaContaFinanceira

  If (viContaFinanceira > 0) Then
    Dim ContaFinanceiraDll As Object
      Set ContaFinanceiraDll = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")
    ContaFinanceiraDll.Exec(CurrentSystem, viContaFinanceira)
  Else
    bsShowMessage("Conta financeira não encontrada", "I")
  End If

End Sub

Public Sub BOTAOMATRICULA_OnClick()
  Dim dllBSInterface0033 As Object
  Dim viHMatricula       As Long

  viHMatricula     = CurrentQuery.FieldByName("MATRICULA").AsInteger

  Set dllBSInterface0033 = CreateBennerObject("BSInterface0033.Matricula")

  dllBSInterface0033.Exec(CurrentSystem, _
                          viHMatricula, _
                          CurrentQuery.FieldByName("CONTRATO").AsInteger, _
                          CurrentQuery.State = 3)

  If CurrentQuery.State <> 1 And viHMatricula > 0 Then

    If CurrentQuery.State = 3 Then
      CurrentQuery.FieldByName("MATRICULA").AsInteger = viHMatricula
    End If

    Dim qSQL As BPesquisa
    Set qSQL = NewQuery

    qSQL.Add("SELECT DATANASCIMENTO,")
    qSQL.Add("       DATACASAMENTO,")
    qSQL.Add("       DATAADOCAO,")
    qSQL.Add("       SEXO")
    qSQL.Add("FROM SAM_MATRICULA")
    qSQL.Add("WHERE HANDLE = :HMATRICULA")
    qSQL.ParamByName("HMATRICULA").AsInteger = viHMatricula
    qSQL.Active = True

    gdDataNascimento = qSQL.FieldByName("DATANASCIMENTO").AsDateTime
    gdDataCasamento  = qSQL.FieldByName("DATACASAMENTO").AsDateTime
    gdDataAdocao     = qSQL.FieldByName("DATAADOCAO").AsDateTime
    gsSexo           = qSQL.FieldByName("SEXO").AsString

  Dim psControleCarencia As String
    If (psControleCarencia = "2") Then
      If (gdDataNascimento > CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime) Then
        CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime = gdDataNascimento
      End If

      If (gdDataAdocao > CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime) Then
        CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime = gdDataAdocao
      End If

      If (gdDataCasamento > CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime) Then
        CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime = gdDataCasamento
      End If
    End If

  If ( CurrentQuery.State = 3) Then
      Dim qFamilia As BPesquisa
      Set qFamilia = NewQuery
      qFamilia.Add("SELECT DATAINCLUSAO, DATAADESAO FROM SAM_FAMILIA WHERE HANDLE = :PHANDLEFAMILIA")
      qFamilia.ParamByName("PHANDLEFAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
      qFamilia.Active = True
      If ( Int(qFamilia.FieldByName("DATAINCLUSAO").AsDateTime) = ServerDate) Then
        CurrentQuery.FieldByName("DATAADESAO").AsDateTime = qFamilia.FieldByName("DATAADESAO").AsDateTime
        CurrentQuery.FieldByName("DATAPRIMEIRAADESAO").AsDateTime = qFamilia.FieldByName("DATAADESAO").AsDateTime
      End If
      Set qFamilia = Nothing
  End If

    Set qSQL = Nothing
  End If

  IDADE.Text = "Idade:" + Str(CalculaIdadeBeneficiario(CurrentQuery.FieldByName("MATRICULA").AsInteger))
  Set dllBSInterface0033 = Nothing

  If CurrentQuery.State <> 3 Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If

End Sub

Public Sub BOTAOPROPAGARENDERECO_OnClick()

  If bsShowMessage("Confirma propagação de endereços do Titular para os dependentes?", "Q") = vbYes Then
    Dim qAtualizaEnd As Object
    Dim qBuscaFilial As Object

    On Error GoTo Finally
      Set qBuscaFilial = NewQuery
      qBuscaFilial.Clear
      qBuscaFilial.Add("SELECT FILIALCUSTO, REGIAO       ")
      qBuscaFilial.Add("      FROM SAM_BENEFICIARIO      ")
      qBuscaFilial.Add("     WHERE FAMILIA = :FAMILIA    ")
      qBuscaFilial.Add("        AND CONTRATO = :CONTRATO ")
      qBuscaFilial.Add("        AND EHTITULAR = 'S'      ")
      qBuscaFilial.ParamByName("FAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
      qBuscaFilial.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
      qBuscaFilial.Active = True

      Set qAtualizaEnd = NewQuery
      qAtualizaEnd.Clear
      qAtualizaEnd.Add("UPDATE SAM_BENEFICIARIO ")
      qAtualizaEnd.Add("   SET ENDERECORESIDENCIAL = Null,    ")
      qAtualizaEnd.Add("       ENDERECOCOMERCIAL = Null,      ")
      qAtualizaEnd.Add("       ENDERECOCORRESPONDENCIA = Null,")
      qAtualizaEnd.Add("       ENDERECOATENDIMENTODOMICILIAR = Null,")
      qAtualizaEnd.Add("     FILIALCUSTO = :FILIALCUSTO,    ")
      qAtualizaEnd.Add("      REGIAO = :REGIAO               ")
      qAtualizaEnd.Add(" WHERE FAMILIA = :FAMILIA             ")
      qAtualizaEnd.Add("  AND CONTRATO = :CONTRATO            ")
      qAtualizaEnd.Add("  AND EHTITULAR = 'N'                 ")
      qAtualizaEnd.ParamByName("FAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
      qAtualizaEnd.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
      qAtualizaEnd.ParamByName("FILIALCUSTO").Value = qBuscaFilial.FieldByName("FILIALCUSTO").AsInteger
      qAtualizaEnd.ParamByName("REGIAO").Value = qBuscaFilial.FieldByName("REGIAO").AsInteger
      qAtualizaEnd.ExecSQL

      Set qAtualizaEnd = Nothing
      Set qBuscaFilial = Nothing

      bsShowMessage("Endereço(s) de "+ CurrentQuery.FieldByName("NOME").AsString +" propagado(s) com sucesso para seus dependentes!", "I")

    Finally:
      Set qAtualizaEnd = Nothing
      Err.Raise(Err.Number, Err.Source, "Falha ao propagar endereços do Beneficiário Titular: " + Err.Description)
  End If

End Sub

Public Sub BOTAOREATIVAR_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
                   1, _
                   "TV_FORM0005", _
                   "Reativação de Beneficiário", _
                   0, _
                   120, _
                   230, _
                   False, _
                   vsMensagem, _
                   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
  Case -1
    bsShowMessage("Operação cancelada pelo usuário!", "I")
  Case 1
    bsShowMessage(vsMensagem, "I")
  End Select

  Set BSINTERFACE0002 = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOREATIVARESPECIAL_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim BSINTERFACE0002 As Object

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

  viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
                   1, _
                   "TV_FORM0007", _
                   "Reativação de Beneficiário", _
                   0, _
                   120, _
                   230, _
                   False, _
                   vsMensagem, _
                   vcContainer)

  Set vcContainer = Nothing

  Select Case viRetorno
  Case -1
    bsShowMessage("Operação cancelada pelo usuário!", "I")
  Case 1
    bsShowMessage(vsMensagem, "I")
  End Select

  Set BSINTERFACE0002 = Nothing

  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub VerificaAdesaoRecemNascido()
  Dim viIdade As Long

  Dim qContrato As BPesquisa
  Set qContrato = NewQuery

  Dim vdllbsben001 As Object
  Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")

  qContrato.Clear
  qContrato.Add("SELECT * FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  qContrato.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qContrato.Active = True

  If Not CurrentQuery.FieldByName("DATAADESAO").IsNull Then
  If (gdDataNascimento <= CurrentQuery.FieldByName("DATAADESAO").AsDateTime) And _
     ((CurrentQuery.FieldByName("DATAADESAO").AsDateTime - gdDataNascimento) <= giPrazoAdesaoRecemNascidosAdotivos) Then
    If VisibleMode Then
      If bsShowMessage("Beneficiário recém-nascido com menos de " + _
                   CStr(giPrazoAdesaoRecemNascidosAdotivos) + _
                   " dias. Isentar Carências?", "Q") = vbYes Then
        If qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then
          CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = gdDataNascimento + giPrazoAdesaoRecemNascidosAdotivos
          CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = 0
          gbIsentouCarenciaRecemNascido = True
        Else
          CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = vdllbsben001.DiasCarencia(CurrentSystem, CurrentQuery.FieldByName("CONTRATO").AsInteger, _
                                                   CurrentQuery.FieldByName("FAMILIA").AsInteger, _
                                                   gsSexo)
        End If
      End If
      DATAPRIMEIRAADESAO.SetFocus
    Else
      bsShowMessage("Beneficiário recém-nascido com menos de " + _
                  CStr(giPrazoAdesaoRecemNascidosAdotivos) + _
                  " dias. Isentou-se suas Carências", "I")

      If qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then
        CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = gdDataNascimento + giPrazoAdesaoRecemNascidosAdotivos
        CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = 0

        gbIsentouCarenciaRecemNascido = True
      Else
        CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = vdllbsben001.DiasCarencia(CurrentSystem, CurrentQuery.FieldByName("CONTRATO").AsInteger, _
                                                 CurrentQuery.FieldByName("FAMILIA").AsInteger, _
                                                 gsSexo)
      End If
    End If
  ElseIf gdDataAdocao > 0 Then
    viIdade = DateDiff("yyyy", _
               CurrentQuery.FieldByName("DATAADESAO").AsDateTime, _
               gdDataNascimento)

    If (viIdade <= 12) And _
       (gdDataAdocao <= CurrentQuery.FieldByName("DATAADESAO").AsDateTime) And _
       ((CurrentQuery.FieldByName("DATAADESAO").AsDateTime - gdDataAdocao) <= giPrazoAdesaoRecemNascidosAdotivos) Then
      If VisibleMode Then
        If bsShowMessage("Beneficiário adotivo com menos de " + _
                 CStr(giPrazoAdesaoRecemNascidosAdotivos) + _
                 " dias. Descontar carências?", "Q") = vbYes Then
          gbDescontouCarenciasAdotivo = True

          If qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then
            CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = gdDataAdocao + giPrazoAdesaoRecemNascidosAdotivos
            CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = 0
          Else
            CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = (CurrentQuery.FieldByName("DATAADESAO").AsDateTime - gdAdesaoTitular) + giDiasCompraCarenciaTitular
          End If
        Else
          gbDescontouCarenciasAdotivo = False
        End If

        DATAPRIMEIRAADESAO.SetFocus
      Else
        bsShowMessage("Beneficiário adotivo com menos de " + _
                CStr(giPrazoAdesaoRecemNascidosAdotivos) + _
                " dias. Descontou-se suas carências", "I")

        gbDescontouCarenciasAdotivo = True

        If qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then
          CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = gdDataAdocao + giPrazoAdesaoRecemNascidosAdotivos
          CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = 0
        Else
          CurrentQuery.FieldByName("DIASCOMPRACARENCIA").AsInteger = (CurrentQuery.FieldByName("DATAADESAO").AsDateTime - gdAdesaoTitular) + giDiasCompraCarenciaTitular
        End If
      End If
    Else
      gbDescontouCarenciasAdotivo = False
    End If
  End If
  End If

  Set qContrato = Nothing
  Set vdllbsben001 = Nothing

End Sub

Public Sub BOTAOVERIFICARINADIMPLENCIA_OnClick()
  Dim vDll As Object
  Dim qContaFinanceira As BPesquisa
  Set vDll = CreateBennerObject("SamContaFinanceira.Consulta")

  Dim handleContaFinanceira As Long
  Dim responsavelconta      As Long
  Set qContaFinanceira = NewQuery

  handleContaFinanceira = RetornaContaFinanceira

  qContaFinanceira.Add(" SELECT * FROM SFN_CONTAFIN WHERE HANDLE = :HANDLE")
  qContaFinanceira.ParamByName("HANDLE").AsInteger = handleContaFinanceira
  qContaFinanceira.Active = True

  If (qContaFinanceira.FieldByName("TABRESPONSAVEL").AsInteger = 1) Then
	responsavelconta = qContaFinanceira.FieldByName("BENEFICIARIO").AsInteger
  Else
    If (qContaFinanceira.FieldByName("TABRESPONSAVEL").AsInteger = 2) Then
  	  responsavelconta = qContaFinanceira.FieldByName("PRESTADOR").AsInteger
    Else
	  responsavelconta = qContaFinanceira.FieldByName("PESSOA").AsInteger
    End If
  End If

  Dim pRetorno As String

  pRetorno = vDll.ConsultaRestricoes(CurrentSystem, qContaFinanceira.FieldByName("TABRESPONSAVEL").AsInteger, responsavelconta)

  If pRetorno <> "" Then
    bsshowmessage(pRetorno, "E")
  End If

  Set vDll = Nothing
  Set qContaFinanceira = Nothing
End Sub

Public Sub DATAADESAO_OnChange()
  Dim qContrato As BPesquisa
  Set qContrato = NewQuery

  qContrato.Clear
  qContrato.Add("SELECT * FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  qContrato.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qContrato.Active = True

  If (CurrentQuery.State = 3) And _
     (CurrentQuery.FieldByName("DATAADESAO").AsDateTime = 0) Then

    If qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then
      CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
    End If

    If qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then
      If gbIsentouCarenciaRecemNascido Then
        CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = gdDataNascimento + giPrazoAdesaoRecemNascidosAdotivos
      End If

      If gbDescontouCarenciasAdotivo Then
        CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = gdDataAdocao + giPrazoAdesaoRecemNascidosAdotivos
      End If
    End If
  End If

  Set qContrato = Nothing
End Sub

Public Sub DATAADESAO_OnExit()
  If CurrentQuery.State = 3 Then
    If  Not CurrentQuery.FieldByName("DATAADESAO").AsDateTime = 0 Then
      CurrentQuery.FieldByName("DATAPRIMEIRAADESAO").AsDateTime = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
      If gsControleCarencia = "2" Then
        CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
      End If
    Else
      CurrentQuery.FieldByName("DATAPRIMEIRAADESAO").Value = Null
    End If
  End If
End Sub

Public Sub DATAADMISSAO_OnChange()
  If (CurrentQuery.State = 3) And _
     (CurrentQuery.FieldByName("DATAADMISSAO").AsDateTime = 0) Then
    If (gsControleCarencia = "2") And _
       (CurrentQuery.FieldByName("DATAADMISSAO").AsDateTime > gdDataAdesaoDireitoPlano) And _
       (CurrentQuery.FieldByName("DATAADMISSAO").AsDateTime > CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime) Then
      CurrentQuery.FieldByName("DATADIREITOPLANO").AsDateTime = CurrentQuery.FieldByName("DATAADMISSAO").AsDateTime
    End If
  End If
End Sub

Public Sub DETALHESBENEFICIARIO_OnClick()
  If (CurrentQuery.FieldByName("HANDLE").AsInteger = 0) Then
    bsShowMessage("É necessário selecionar um beneficiário", "I")

    Exit Sub
  End If

  If (CurrentQuery.State = 2) Or _
     (CurrentQuery.State = 3) Then
    bsShowMessage("O registro não pode estar em edição!", "I")

    Exit Sub
  End If

  Dim dllCA006 As Object

  Set dllCA006 = CreateBennerObject("CA006.ConsultaBeneficiario")

  dllCA006.Info(CurrentSystem, _
            CurrentQuery.FieldByName("HANDLE").AsString, _
            0)

    Set dllCA006 = Nothing
End Sub

Public Sub DIGITAR_OnClick()
  Dim dllBSInterface001 As Object

  Set dllBSInterface001 = CreateBennerObject("BSINTERFACE0011.DigitarBeneficiario")

  dllBSInterface001.Exec(CurrentSystem, _
             CurrentQuery.FieldByName("HANDLE").AsInteger, _
             CurrentQuery.FieldByName("CONTRATO").AsInteger, _
             CurrentQuery.FieldByName("FAMILIA").AsInteger)

  Set dllBSInterface001 = Nothing
  If VisibleMode Then
    RefreshNodesWithTable("")
  End If
End Sub

Public Sub FAMILIA_OnChange()
  Dim query As BPesquisa
  Set query = NewQuery

  query.Clear
  query.Add("SELECT C.TABTIPOCONTRATO, F.DATAADESAO")
  query.Add("  FROM SAM_FAMILIA F")
  query.Add("  JOIN SAM_CONTRATO C ON C.HANDLE = F.CONTRATO")
  query.Add(" WHERE F.HANDLE = :FAMILIA")
  query.ParamByName("FAMILIA").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  query.Active = True

  If query.FieldByName("TABTIPOCONTRATO").AsInteger <> 1 Then
      gdDataAdesaoDireitoPlano = query.FieldByName("DATAADESAO").AsDateTime
    End If
End Sub

Public Sub MATRICULAINDICADORA_OnChange()
  If CurrentQuery.FieldByName("EHTITULAR").AsString = "S" Then
      giMatriculaIndicadora = CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger
    End If
End Sub

Public Sub MATRICULAINDICADORA_OnExit()

  If CurrentQuery.State = 1 Then
    Exit Sub
  End If

  Dim qAux As BPesquisa

  Set qAux = NewQuery

  giIndiceBusca = 0

  qAux.Active = False
  qAux.Clear

  qAux.Add("SELECT HANDLE,    ")
  qAux.Add("     NOME      ")
  qAux.Add("  FROM SAM_MATRICULA")
  qAux.Add("WHERE HANDLE = :MATRICULA")

  qAux.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger

  If (Not ValidaNumero(MATRICULAINDICADORA.Text)) Then

    gsMatriculaDigitada = MATRICULAINDICADORA.Text
    giIndiceBusca       = 2
  End If

  qAux.Add("   AND EHINDICADOR = 'S'")

  qAux.Active = True

  If qAux.EOF Then
    giMatriculaIndicadora = 0

    MATRICULAINDICADORA_OnPopup(False)

    If giMatriculaIndicadora <= 0 Then CurrentQuery.FieldByName("MATRICULAINDICADORA").Clear
  Else
    CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger = qAux.FieldByName("HANDLE").AsInteger
  End If

  qAux.Active = False

  qAux.Clear

  qAux.Add("SELECT HANDLE,         ")
  qAux.Add("     NOME           ")
  qAux.Add("  FROM SAM_BENEFICIARIO     ")
  qAux.Add(" WHERE MATRICULA = :MATRICULA")

  qAux.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger

  qAux.Active = True

  If Not qAux.EOF Then CurrentQuery.FieldByName("BENEFICIARIOINDICADOR").AsInteger = qAux.FieldByName("HANDLE").AsInteger

  giMatriculaIndicadora = CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger
  'gsMatriculaIndicadora = ""
  gsMatriculaDigitada   = ""
  giIndiceBusca         = 0
End Sub

Public Sub MATRICULAINDICADORA_OnPopup(ShowPopup As Boolean)
  Dim vsColunas As String
  Dim vsWhere As String
  Dim viIndiceBusca As Long
  Dim viHandle As Long
  Dim dllProcura As Object

  ShowPopup = False

  vsColunas = "MATRICULA|Z_NOME"
  vsWhere = "EHINDICADOR = 'S'"

  If (giIndiceBusca > 0) Then
    viIndiceBusca = giIndiceBusca
  Else
    viIndiceBusca = 0
  End If

  Set dllProcura = CreateBennerObject("PROCURA.Procurar")

  viHandle = dllProcura.Exec(CurrentSystem, _
                 "SAM_MATRICULA", _
                 vsColunas, _
                 1, _
                 "Matrícula|Nome", _
                 vsWhere, _
                 "Procura por matrícula", _
                 False, _
                 "")

  Set dllProcura = Nothing

  giMatriculaIndicadora = viHandle

  If (viHandle <> 0) Then CurrentQuery.FieldByName("MATRICULAINDICADORA").AsInteger = viHandle
End Sub

Public Sub MOTIVOBLOQUEIO_OnPopup(ShowPopup As Boolean)
  Dim qContrato As BPesquisa

  Set qContrato = NewQuery

  qContrato.Add("SELECT TABADESAORECEBIMENTO, MOTIVOBLOQUEIOAUTOMATICO")
  qContrato.Add("  FROM SAM_CONTRATO      ")
  qContrato.Add(" WHERE HANDLE = :CONTRATO  ")

  qContrato.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger

  qContrato.Active = True

  If (qContrato.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2) Then
    MOTIVOBLOQUEIO.LocalWhere = "SAM_MOTIVOBLOQUEIO.HANDLE <> " + qContrato.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").AsString
  End If

  Set qContrato = Nothing
End Sub

Public Sub TABLE_AfterCancel()
  DIGITAR.Visible               = True
  BOTAOBENEFICIARIO.Visible     = True
  BOTAOCANCELAR.Visible         = True
  BOTAOCARTAO.Visible           = True
  BOTAOFINANCEIRO.Visible       = True
  BOTAOREATIVAR.Visible         = True
  BOTAOREATIVARESPECIAL.Visible = True
  DETALHESBENEFICIARIO.Visible  = True
  'BOTAOENDERECO.Visible         = False

  If InTransaction Then
    Rollback
  End If

  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""

  CarregaRotulosEndereco

End Sub

Public Sub TABLE_AfterCommitted()
  If gbGerarCartao And _
     VisibleMode Then

     On Error GoTo erro:

       Dim dllBSInterface0002 As Object
       Dim vcContainer        As CSDContainer
       Dim vsMensagem         As String

       Set dllBSInterface0002 = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
       Set vcContainer        = NewContainer

     If Not InTransaction Then
          StartTransaction
       End If

       dllBSInterface0002.Exec(CurrentSystem, _
                               1, _
                               "TV_FORM0013", _
                               "Geração do cartão", _
                               0, _
                               180, _
                               500, _
                               False, _
                               vsMensagem, _
                               vcContainer)

       Set vcContainer        = Nothing
       Set dllBSInterface0002 = Nothing

       If vsMensagem <> "" Then
         bsShowMessage(vsMensagem, "I")
       End If

     If InTransaction Then
           Commit
       End If

       Exit Sub
     erro:
       If InTransaction Then
         Rollback
       End If
       Set vcContainer        = Nothing
       Set dllBSInterface0002 = Nothing
  End If
End Sub

Public Sub TABLE_AfterEdit()

End Sub

Public Sub TABLE_AfterInsert()
  Dim viCodigoBenefAns As Long
  Dim vsMensagem As String
  Dim query As BPesquisa
  Set query = NewQuery

  query.Clear
  query.Add("SELECT NUMEROBENEFAUTOMATICO FROM SAM_FAMILIA WHERE HANDLE = :HANDLE")
  query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger
  query.Active = True

  If query.FieldByName("NUMEROBENEFAUTOMATICO").AsString = "N" Then
    CODIGODEPENDENTE.ReadOnly = False
  Else
    CODIGODEPENDENTE.ReadOnly = True
  End If

  query.Active = False
  query.Clear
  query.Add("SELECT TABADESAORECEBIMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  query.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  query.Active = True

  If query.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2 Then MOTIVOBLOQUEIO.ReadOnly = True

  Dim vdllbsben001 As Object
  Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")

  vdllbsben001.InserirBeneficiario(CurrentSystem, _
                 CurrentQuery.TQuery, _
                 gsControleCarencia, _
                 gsTitularAutonomo, _
                 gdDataAdmissaoTitular, _
                 gdDataAdesaoDireitoPlano, _
                 gdAdesaoTitular, _
                 giPrazoAdesaoRecemNascidosAdotivos, _
                 giDiasCompraCarenciaTitular, _
                 giOrigemCarenciaTitular, _
                 vsMensagem)

  If vsMensagem <> "" Then bsShowMessage(vsMensagem, "I")

  Set vdllbsben001 = Nothing

  ATENDIMENTOATE.ReadOnly = True
End Sub

Public Sub TABLE_AfterPost()

  If Not CurrentQuery.IsVirtual Then
    Dim vsMensagem As String
    Dim viRetorno  As Long
    Dim qQuery     As Object

    Set qQuery = NewQuery

	If (WebMode And CurrentEntity.TransitoryVars("MODULOSBENEFICIARIODIGITACAO").IsPresent) Then
	  Dim vsModulosHandles As String
	  Dim vdllBsBen022 As Object

	  vsModulosHandles = CurrentEntity.TransitoryVars("MODULOSBENEFICIARIODIGITACAO").AsString
      Set vdllBsBen022 = CreateBennerObject("BSBEN022.Modulo")

      viRetorno = vdllBsBen022.IncluirModulosWeb(CurrentSystem, _
	                  							 CurrentQuery.TQuery, _
	                  							 vsModulosHandles, _
	                  							 vsMensagem)

	  Set vdllBsBen022 = Nothing

	  If viRetorno = 1 Then
        Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Gravação cancelada!")
      End If
    End If

    qQuery.Add("SELECT BENCRITICAEMITECARTAO ")
    qQuery.Add("  FROM SAM_CONTRATO       ")
    qQuery.Add(" WHERE (HANDLE = :HANDLE)     ")
    qQuery.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
      qQuery.Active = True

    Dim vdllbsben001 As Object
      Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")

    viRetorno = vdllbsben001.AfterPost(CurrentSystem, _
                         CurrentQuery.TQuery, _
                         vsModoEdicao, _
                         (qQuery.FieldByName("BENCRITICAEMITECARTAO").AsString = "S"), _
                         giDependente, _
                         giHTipoDependenteAnterior, _
                         giTitularConjuge, _
                         gsCodigoTabelaPrcTitularAnterior, _
                         gsGrupoDependenteAnterior, _
                         gbGerarCartao, _
                         vsMensagem)

    Set qQuery       = Nothing
    Set vdllbsben001 = Nothing

    If viRetorno = 1 Then
        Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Gravação cancelada!")
    Else
      If vsMensagem <> "" Then
        bsShowMessage(vsMensagem, "I")
      End If


     End If
    End If

  RegistrarLogAlteracao "SAM_BENEFICIARIO", CurrentQuery.FieldByName("HANDLE").AsInteger, "MACRO TABLE_AfterPost"

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If

  If vsModoEdicao = "A" Then

    Dim dllBSBen021 As Object
    Set dllBSBen021 = CreateBennerObject("BSBen021.AtualizacaoEndereco")

    viRetorno = dllBSBen021.Excluir(CurrentSystem, _
                                    vsXMLEnderecosExcluidos, _
                                    vsMensagem)

    Set dllBSBen021 = Nothing

    If viRetorno = 1 Then
       Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Gravação cancelada!")
    Else
       If vsMensagem <> "" Then
         bsShowMessage(vsMensagem, "I")
       End If
    End If
  End If

    'Se estiver sendo executado através dos componentes virtuais a transação deve ser fechada "manualmente"
    'Este procedimento é necessário pois os componentes de engine virtual não controlam transação
    If CurrentQuery.IsVirtual And _
       VisibleMode And _
       InTransaction Then
      Commit
    End If

  If (VisibleMode And _
     (Not CurrentQuery.IsVirtual)) Then
    UpdateLastUpdate("SAM_ENDERECO")
    SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If
  'limpa as variáveis de endereço para que ao abrir a tela novamente
  'pegue as informações da base de dados. deixando pronto para outra alteração.
  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""

  Dim TQIntegracoesCorpBennerBLL As CSBusinessComponent
  Set TQIntegracoesCorpBennerBLL = BusinessComponent.CreateInstance("Benner.Saude.IntegracaoFinanceira.Business.TabelasBasicas.TQIntegracoesCorpBennerBLL, Benner.Saude.IntegracaoFinanceira.Business")

  TQIntegracoesCorpBennerBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "SAM_BENEFICIARIO")
  TQIntegracoesCorpBennerBLL.AddParameter(pdtString, "X")

  TQIntegracoesCorpBennerBLL.Execute("InserirDadosIntegracao")
End Sub

Public Sub TABLE_AfterScroll()

	Dim mascaraTelefone As String
	Dim vEspecificoDLL As Object
  	Set vEspecificoDLL = CreateBennerObject("Especifico.uEspecifico")
	mascaraTelefone = vEspecificoDLL.BEN_MascaraTelefone(CurrentSystem)

	CurrentQuery.FieldByName("CELULAR").Mask = mascaraTelefone
	CurrentQuery.FieldByName("PAGERCENTRAL").Mask = mascaraTelefone

	Set vEspecificoDLL = Nothing

    ATENDIMENTOATE.ReadOnly = True
    ROTRESPONSAVELLEGAL.Text = RetornaResponsavelLegal

    If CurrentQuery.State <> 3 Then
      IDADE.Text = "Idade:" + Str(CalculaIdadeBeneficiario(CurrentQuery.FieldByName("MATRICULA").AsInteger))
    End If

  If WebMode Then
    Dim qSexo As Object
    Set qSexo = NewQuery
    qSexo.Add("SELECT M.SEXO FROM SAM_MATRICULA M JOIN SAM_BENEFICIARIO B ON (B.MATRICULA = M.HANDLE) WHERE B.HANDLE = :HANDLE")
    qSexo.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSexo.Active = True

    If qSexo.FieldByName("SEXO").AsString = "M" Then
      ROTSEXO.Text = "Sexo: Masculino"
    ElseIf qSexo.FieldByName("SEXO").AsString = "F" Then
      ROTSEXO.Text = "Sexo: Feminino"
    End If

        SessionVar("HCONTAFINANCEIRA") = CStr(RetornaContaFinanceira)

  ElseIf VisibleMode Then
    ROTSEXO.Text = ""
  End If


  If CurrentQuery.FieldByName("EHTITULAR").AsString = "N" Then
    BOTAOPROPAGARENDERECO.Enabled = False
  Else
    BOTAOPROPAGARENDERECO.Enabled = True
  End If

    If CurrentQuery.State <> 1 Then
      DIGITAR.Visible               = False
      BOTAOBENEFICIARIO.Visible     = False
      BOTAOCANCELAR.Visible         = False
      BOTAOCARTAO.Visible           = False
      BOTAOFINANCEIRO.Visible       = False
      BOTAOREATIVAR.Visible         = False
      BOTAOREATIVARESPECIAL.Visible = False
      DETALHESBENEFICIARIO.Visible  = False
    BOTAOMIGRAR.Visible                = False
      BOTAOALERTAENDCORRESP.Visible      = False
      BOTAOATUALIZAADESAO.Visible        = False
      'BOTAOENDERECO.Visible         = True

      If WebMode Then
        If CurrentQuery.State = 3 Then
          MATRICULA.ReadOnly          = False
        Else
          MATRICULA.ReadOnly          = True
        End If
      Else
        MATRICULA.ReadOnly            = True
      End If
    Else
      DIGITAR.Visible               = True
      BOTAOBENEFICIARIO.Visible     = True
      BOTAOCANCELAR.Visible         = True
      BOTAOCARTAO.Visible           = True
      BOTAOFINANCEIRO.Visible       = True
      BOTAOREATIVAR.Visible         = True
      BOTAOREATIVARESPECIAL.Visible = True
      DETALHESBENEFICIARIO.Visible  = True
    BOTAOMIGRAR.Visible                = True
      BOTAOALERTAENDCORRESP.Visible      = True
      BOTAOATUALIZAADESAO.Visible        = True

      'BOTAOENDERECO.Visible         = False

      MATRICULA.ReadOnly            = True
    End If
  BOTAOENDERECO.Visible         = True
  BOTAOALTERACONTAFINANCEIRA.Visible = False

  SessionVar("piDiasCompraCarenciaTitular")      = CStr(giDiasCompraCarenciaTitular)
  SessionVar("piOrigemCarenciaTitular")         = CStr(giOrigemCarenciaTitular)
  SessionVar("pdAdesaoTitular")             = CStr(gdAdesaoTitular)
  SessionVar("pdDataAdmissaoTitular")           = CStr(gdDataAdmissaoTitular)
  SessionVar("psTitularAutonomo")             = gsTitularAutonomo
  SessionVar("HBENEFICIARIO")                    = CurrentQuery.FieldByName("HANDLE").AsString
  SessionVar("EMISSAOCARTAO_ALTERACAOCADASTRAL") = "N"

  Dim vdDataFinalSuspensao As Date
  Dim vsBeneficiario As String
  Dim vsMascaraBeneficiario As String
  Dim vsCodigoFormatado As String
  Dim query As BPesquisa
  Dim qOdonto As Object
  Dim qParametrosBeneficiario As BPesquisa

  Dim qContratoPadraoCampos As BPesquisa

  Dim vdllbsben001 As Object
  Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")
    vdllbsben001.Inicializar(CurrentSystem)

  Set qParametrosBeneficiario = NewQuery
  Set query = NewQuery
  Set qContratoPadraoCampos = NewQuery
    Set qOdonto = NewQuery

    qOdonto.Active = False
    qOdonto.Clear
    qOdonto.Add("SELECT REGODONTOINDEPENDENTE FROM SAM_PARAMETROSANS")
    qOdonto.Active = True

    If qOdonto.FieldByName("REGODONTOINDEPENDENTE").AsString = "N" Then
       CCOODONTO.Visible       = False
       CCODVODONTO.Visible     = False
       CODIGOANSODONTO.Visible = False
    Else
       CCOODONTO.Visible       = True
       CCODVODONTO.Visible     = True
       CODIGOANSODONTO.Visible = True
    End If

    Set qOdonto = Nothing


  qParametrosBeneficiario.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")

  qParametrosBeneficiario.Active = True

  CarregaRotulosEndereco

  'If (CurrentQuery.State = 3) Then
  '  gbInclusaoDeRegistroDoBeneficiario = True
  'Else
  '  gbInclusaoDeRegistroDoBeneficiario = False
  'End If

  PreparaNumeracaoBenef

  If vdllbsben001.VerificaSuspensao(CurrentSystem, _
                  CurrentQuery.FieldByName("HANDLE").AsInteger, _
                  CurrentQuery.FieldByName("FAMILIA").AsInteger, _
                  CurrentQuery.FieldByName("CONTRATO").AsInteger, _
                  vdDataFinalSuspensao) Then
    BOTAOCANCELAR.Enabled       = False
    BOTAOREATIVAR.Enabled       = False
    BOTAOREATIVARESPECIAL.Enabled = False
    BOTAOCARTAO.Enabled       = False
    BOTAOMIGRAR.Enabled       = False

    If (vdDataFinalSuspensao = 0) Then
      SUSPENSO.Text = "Suspenso indefinidamente"
    Else
      SUSPENSO.Text = "Suspenso até " + CStr(vdDataFinalSuspensao)
    End If
  Else
    SUSPENSO.Text = ""
    BOTAOCANCELAR.Enabled       = True
    BOTAOREATIVAR.Enabled       = True
    BOTAOREATIVARESPECIAL.Enabled = True
    BOTAOCARTAO.Enabled       = True
    BOTAOMIGRAR.Enabled       = True
  End If

  VerificaLicenca

  If (CurrentQuery.FieldByName("EHTITULAR").AsString = "S") And _
     (CurrentQuery.State = 2) Then
    EHTITULAR.ReadOnly = True
  Else
    EHTITULAR.ReadOnly = False
  End If

  If (CurrentQuery.State = 3) Then
    qContratoPadraoCampos.Add("SELECT BENPERMITEADIANTAMENTO,   ")
    qContratoPadraoCampos.Add("      BENENDERECOCORRESPONDENCIA,")
    qContratoPadraoCampos.Add("      BENCRITICAEMITECARTAO     ")
    qContratoPadraoCampos.Add("   FROM SAM_CONTRATO         ")
    qContratoPadraoCampos.Add("  WHERE (HANDLE = :HANDLE)     ")

    qContratoPadraoCampos.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger

    qContratoPadraoCampos.Active = True

    CurrentQuery.FieldByName("PERMITEADIANTAMENTO").AsString = qContratoPadraoCampos.FieldByName("BENPERMITEADIANTAMENTO").AsString

    gsEnderecoCorrespondencia = qContratoPadraoCampos.FieldByName("BENENDERECOCORRESPONDENCIA").AsString
    gbCriticaEmissaoCartao    = (qContratoPadraoCampos.FieldByName("BENCRITICAEMITECARTAO").AsString = "S")

    qContratoPadraoCampos.Active = False
  End If

  CarregaRotulosEndereco

  'If vFlagMascara = 0 Then
  '  vsBeneficiario = CurrentQuery.FieldByName("BENEFICIARIO").AsString
  '  viFlagMascara  = 1

  '  Dim dllSamBeneficiario As Object
  '  Set dllSamBeneficiario = CreateBennerObject("SAMBENEFICIARIO.Cadastro")
  '  dllSamBeneficiario.Mascara(CurrentSystem, _
  '                 vsBeneficiario, _
  '                 vsMascaraBeneficiario, _
  '                 vsCodigoFormatado)
  '  Set dllSamBeneficiario = Nothing
  'End If

  'CurrentQuery.FieldByName("BENEFICIARIO").Mask = vsMascaraBeneficiario

  query.Active = False

  query.Clear
  query.Add("SELECT TABTIPOGESTAO")
  query.Add("   FROM EMPRESAS      ")
  query.Add("  WHERE HANDLE =  :E ")

  query.ParamByName("E").AsInteger = CurrentQuery.FieldByName("EMPRESA").AsInteger

  query.Active = True

  If (query.FieldByName("TABTIPOGESTAO").AsInteger = 3) Then
    GRUPOCOOP.Visible = True
  Else
    GRUPOCOOP.Visible = False
  End If

    vdllbsben001.Finalizar
  Set vdllbsben001 = Nothing

  gsSituacaoRhAux = CurrentQuery.FieldByName("SITUACAORH").AsString
  gsDataSituacaoRh = CurrentQuery.FieldByName("DATAULTIMAALTERACAOSITUACAORH").AsString

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
    'Se estiver em modo desktop a transação deve ser iniciada antes da edição
    'pela possibilidade de inclusão/alteração de endereços
    If VisibleMode Then
      If Not InTransaction Then
        StartTransaction
      End If
    End If

    DIGITAR.Visible               = False
    BOTAOBENEFICIARIO.Visible     = False
    BOTAOCANCELAR.Visible         = False
    BOTAOCARTAO.Visible           = False
    BOTAOFINANCEIRO.Visible       = False
    BOTAOREATIVAR.Visible         = False
    BOTAOREATIVARESPECIAL.Visible = False
    DETALHESBENEFICIARIO.Visible  = False
    'BOTAOENDERECO.Visible         = True

    MATRICULA.ReadOnly            = True




    SessionVar("EMISSAOCARTAO_ALTERACAOCADASTRAL") = "S"

  Dim vsMensagem As String
  Dim query As BPesquisa
  Set query = NewQuery

  vsModoEdicao = "A"

  giHTipoDependenteAnterior = CurrentQuery.FieldByName("TIPODEPENDENTE").AsInteger
  giHEstadoCivilAnterior    = CurrentQuery.FieldByName("ESTADOCIVIL").AsInteger
  gbGerarCartao            = False

  query.Clear
  query.Add("SELECT T.GRUPODEPENDENTE       ")
  query.Add("  FROM SAM_TIPODEPENDENTE T,     ")
  query.Add("      SAM_CONTRATO_TPDEP C     ")
  query.Add("  WHERE C.HANDLE = :TIPODEPENDENTE ")
  query.Add("    AND C.TIPODEPENDENTE = T.HANDLE")
  query.ParamByName("TIPODEPENDENTE").AsInteger = giHTipoDependenteAnterior
  query.Active = True

  gsGrupoDependenteAnterior = query.FieldByName("GRUPODEPENDENTE").AsString

  giTitularConjuge = CurrentQuery.FieldByName("TITULARCONJUGE").AsInteger
  Dim vdllbsben001 As Object
  Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")

  vdllbsben001.EditarBeneficiario(CurrentSystem, _
                CurrentQuery.TQuery, _
                gsNaoFaturarModulosAnterior, _
                gsNaoFaturarGuiasAnterior, _
                gsNaoTemCarenciaAnterior, _
                gsCodigoTabelaPrcTitularAnterior, _
                giDiasCompraCarenciaAnterior, _
                giMotivoBloqueioBenef, _
                gbAlteracaoCadastral, _
                vsMensagem)

  Set vdllbsben001 = Nothing

  If vsMensagem <> "" Then
    bsShowMessage(vsMensagem, "I")

      ATENDIMENTOATE.ReadOnly = False
      BENEFICIARIOINDICADOR.ReadOnly = True
      BENEFICIARIOMIGRACAO.ReadOnly = True
      CARGO.ReadOnly = True
      CARTAOIDENTIFICACAO.ReadOnly = True
      CBO.ReadOnly = True
      CELULAR.ReadOnly = True
      CODIGOANTIGO.ReadOnly = True
      CODIGODEAFINIDADE.ReadOnly = True
      CODIGODEORIGEM.ReadOnly = True
      CODIGOTABELAPRC.ReadOnly = True
      DATAADMISSAO.ReadOnly = True
      DATADIREITOPLANO.ReadOnly = True
      DATAINICIOAPOSENTADORIA.ReadOnly = True
      DEMONSTRATIVOFINANCEIROINDIVID.ReadOnly = True
      DESTINOCOBCONTRIBUICAOSOCIAL.ReadOnly = True
      DESTINOCOBFRQURGEMERG.ReadOnly = True
      DESTINOCOBMENSALIDADE.ReadOnly = True
      DESTINOCOBPFSERVICO.ReadOnly = True
      DIASCOMPRACARENCIA.ReadOnly = True
      DIREITOAADIANTAMENTO.ReadOnly = True
      EHTITULAR.ReadOnly = True
      EMAIL.ReadOnly = True
      ESFABRANGENCIA.ReadOnly = True
      ESFEQUIPE.ReadOnly = True
      ESTADOCIVIL.ReadOnly = True
      IDENTIFICADORCARTAO.ReadOnly = True
      INFORMATIVOS.ReadOnly = True
      LIMINAR.ReadOnly = False
      LIMITEADIANTAMENTO.ReadOnly = True
      LOCATEND.ReadOnly = True
      LOCCOB.ReadOnly = True
      MATRICULAFUNCIONAL.ReadOnly = True
      MATRICULAINDICADORA.ReadOnly = True
      MOTIVOBLOQUEIO.ReadOnly = True
      MOTIVOINCLUSAO.ReadOnly = True
      NAOATIVAREMPRESTIMO.ReadOnly = True
      NAOFATURARGUIAS.ReadOnly = True
      NAOFATURARMODULOS.ReadOnly = True
      NAOPERMITEAUXILIO.ReadOnly = True
      NAOTEMCARENCIA.ReadOnly = True
      NUMEROINSS.ReadOnly = True
      PAGER.ReadOnly = True
      PAGERCENTRAL.ReadOnly = True
      PERMITEADIANTAMENTO.ReadOnly = True
      PERMITEREEMBOLSO.ReadOnly = True
      PRIORITARIO.ReadOnly = True
      PROCURADOR.ReadOnly = True
      PROPORCIONALIDADEACOBRARCANC.ReadOnly = True
      SETOR.ReadOnly = True
    '  TITULARCONJUGE.ReadOnly = True  SMS 101392 - Paulo Melo - 20/08/2008 - Campo pode ser editável em beneficiários ativos e cancelados
      TIPODEPENDENTE.ReadOnly = True
      SITUACAORH.ReadOnly = True
      ORIGEMCARENCIA.ReadOnly = True
      NIVEL.ReadOnly = True
  Else
      ATENDIMENTOATE.ReadOnly = True
      BENEFICIARIOINDICADOR.ReadOnly = False
      BENEFICIARIOMIGRACAO.ReadOnly = False
      CARGO.ReadOnly = False
      CARTAOIDENTIFICACAO.ReadOnly = False
      CBO.ReadOnly = False
      CELULAR.ReadOnly = False
      CODIGOANTIGO.ReadOnly = False
      CODIGODEAFINIDADE.ReadOnly = False
      CODIGODEORIGEM.ReadOnly = False
      CODIGOTABELAPRC.ReadOnly = False
      DATAADMISSAO.ReadOnly = False
      DATADIREITOPLANO.ReadOnly = False
      DATAINICIOAPOSENTADORIA.ReadOnly = False
      DEMONSTRATIVOFINANCEIROINDIVID.ReadOnly = False
      DESTINOCOBCONTRIBUICAOSOCIAL.ReadOnly = False
      DESTINOCOBFRQURGEMERG.ReadOnly = False
      DESTINOCOBMENSALIDADE.ReadOnly = False
      DESTINOCOBPFSERVICO.ReadOnly = False
      DIASCOMPRACARENCIA.ReadOnly = False
      DIREITOAADIANTAMENTO.ReadOnly = False
      EHTITULAR.ReadOnly = False
      EMAIL.ReadOnly = False
      ESFABRANGENCIA.ReadOnly = False
      ESFEQUIPE.ReadOnly = False
      ESTADOCIVIL.ReadOnly = False
      IDENTIFICADORCARTAO.ReadOnly = False
      INFORMATIVOS.ReadOnly = False
      LIMINAR.ReadOnly = False
      LIMITEADIANTAMENTO.ReadOnly = False
      LOCATEND.ReadOnly = False
      LOCCOB.ReadOnly = False
      MATRICULAFUNCIONAL.ReadOnly = False
      MATRICULAINDICADORA.ReadOnly = False
      MOTIVOBLOQUEIO.ReadOnly = False
      MOTIVOINCLUSAO.ReadOnly = False
      NAOATIVAREMPRESTIMO.ReadOnly = False
      NAOFATURARGUIAS.ReadOnly = False
      NAOFATURARMODULOS.ReadOnly = False
      NAOPERMITEAUXILIO.ReadOnly = False
      NAOTEMCARENCIA.ReadOnly = False
      NUMEROINSS.ReadOnly = False
      PAGER.ReadOnly = False
      PAGERCENTRAL.ReadOnly = False
      PERMITEADIANTAMENTO.ReadOnly = False
      PERMITEREEMBOLSO.ReadOnly = False
      PRIORITARIO.ReadOnly = False
      PROCURADOR.ReadOnly = False
      PROPORCIONALIDADEACOBRARCANC.ReadOnly = False
      SETOR.ReadOnly = False
      TITULARCONJUGE.ReadOnly = False
      TIPODEPENDENTE.ReadOnly = False
      SITUACAORH.ReadOnly = False
      ORIGEMCARENCIA.ReadOnly = False
      NIVEL.ReadOnly = False
  End If

    vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  DIGITAR.Visible               = False
  BOTAOBENEFICIARIO.Visible     = False
  BOTAOCANCELAR.Visible         = False
  BOTAOCARTAO.Visible           = False
  BOTAOFINANCEIRO.Visible       = False
  BOTAOREATIVAR.Visible         = False
  BOTAOREATIVARESPECIAL.Visible = False
  DETALHESBENEFICIARIO.Visible  = False
  'BOTAOENDERECO.Visible         = True
  vsModoEdicao = "I"

  If WebMode Then
    MATRICULA.ReadOnly          = False
  Else
    MATRICULA.ReadOnly          = True
  End If

  SessionVar("EMISSAOCARTAO_ALTERACAOCADASTRAL") = "N"

  gbGerarCartao = False

  'Se estiver em modo desktop a transação deve ser iniciada antes da edição
  'pela possibilidade de inclusão/alteração de endereços
  If VisibleMode Then
    If Not InTransaction Then
      StartTransaction
  End If
  End If

  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  VerificaAdesaoRecemNascido

  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vdllbsben001 As Object
  Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")
  Dim qSqlPlano As BPesquisa
    Set qSqlPlano = NewQuery

  viRetorno = vdllbsben001.BeforePost(CurrentSystem, _
                  CurrentQuery.TQuery, _
                  gbGerarCartao, _
                  gbAchouCodigoTabelaPrcTitular, _
                  gbIsentouCarenciaRecemNascido, _
                  gbDescontouCarenciasAdotivo, _
                  gbCriticaEmissaoCartao, _
                  gdDataFechamento, _
                  gdDataNascimento, _
                  gdDataAdocao, _
                  gdDataAdmissaoTitular, _
                  gdAdesaoTitular, _
                  gsCodigoTabelaPrcTitular, _
                  gsNaoFaturarModulosAnterior, _
                  gsNaoFaturarGuiasAnterior, _
                  gsVerificarContribPrevidencia, _
                  gsTitularAutonomo, _
                  gsNaoTemCarenciaAnterior, _
                  gsGrupoDependenteAnterior, _
                  gsCodigoTabelaPrcTitularAnterior, _
                  giHTipoDependenteAnterior, _
                  giHEstadoCivilAnterior, _
                  giPrazoAdesaoRecemNascidosAdotivos, _
                  giMotivoBloqueioBenef, _
                  giDependente, _
                  giOrigemCarenciaTitular, _
                  giDiasCompraCarenciaTitular, _
                  giTitularConjuge, _
                  giUsuarioLiberou, _
                  vsMensagem, _
                  giDiasCompraCarenciaAnterior)

  Set vdllbsben001 = Nothing

  If viRetorno = 2 Then
    If vsMensagem <> "" Then
      If bsShowMessage(vsMensagem + Chr(13) + "Deseja salvar o registro?", "Q") = vbNo Then
        If (Not WebMode) Then
          CanContinue = False
        End If
      End If
    End If

  ElseIf viRetorno = 1 Then
    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "E")
    End If
    CanContinue = False

  Else
    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "E")
    End If
  End If

  If (WebMode And CurrentEntity.TransitoryVars("MODULOSBENEFICIARIODIGITACAO").IsPresent) Then
    Dim vsModulosHandles As String
	Dim vdllBsBen022 As Object

    vsModulosHandles = CurrentEntity.TransitoryVars("MODULOSBENEFICIARIODIGITACAO").AsString
	Set vdllBsBen022 = CreateBennerObject("BSBEN022.Modulo")

  	viRetorno = vdllBsBen022.VerificaModulosWeb(CurrentSystem, _
                  CurrentQuery.TQuery, _
                  vsModulosHandles, _
                  vsMensagem)

    Set vdllBsBen022 = Nothing

    If vsMensagem <> "" Then
      bsShowMessage(vsMensagem, "E")
      CanContinue = False
    End If

  End If

  If ((gsSituacaoRhAux <> CurrentQuery.FieldByName("SITUACAORH").AsString) And (gsDataSituacaoRh = CurrentQuery.FieldByName("DATAULTIMAALTERACAOSITUACAORH").AsString)) Then
    CurrentQuery.FieldByName("DATAULTIMAALTERACAOSITUACAORH").AsDateTime = CurrentSystem.ServerDate
  End If


End Sub

Public Sub CarregaRotulosEndereco
  Dim vcRotulos As CSDContainer
  Dim viHEnderecoResidencial, viHEnderecoComercial, viHEnderecoCorrespondencia, viHEnderecoAtendimentoDomi As Long

  Set vcRotulos = NewContainer

  vcRotulos.AddFields("ROTULORES1: STRING; ROTULORES2: STRING; ROTULORES3: STRING; ROTULORES4: STRING; ROTULORES5: STRING")
  vcRotulos.AddFields("ROTULOCOM1: STRING; ROTULOCOM2: STRING; ROTULOCOM3: STRING; ROTULOCOM4: STRING; ROTULOCOM5: STRING")
  vcRotulos.AddFields("ROTULOCORRESP1: STRING; ROTULOCORRESP2: STRING; ROTULOCORRESP3: STRING; ROTULOCORRESP4: STRING; ROTULOCORRESP5: STRING")
  vcRotulos.AddFields("ROTULOATENDDOM1: STRING; ROTULOATENDDOM2: STRING; ROTULOATENDDOM3: STRING; ROTULOATENDDOM4: STRING; ROTULOATENDDOM5: STRING")

  vcRotulos.Insert
  Dim vdllbsben001 As Object
  Set vdllbsben001 = CreateBennerObject("BSBEN001.Beneficiario")


  viHEnderecoResidencial = CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger
  viHEnderecoComercial = CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger
  viHEnderecoCorrespondencia = CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger
  viHEnderecoAtendimentoDomi = CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger

  'Verificando se beneficiário não possui endereço, se não é titular e adicionando o endereço do titular da família
  If (viHEnderecoResidencial = 0) Or (viHEnderecoComercial = 0) Or (viHEnderecoCorrespondencia = 0) Or (viHEnderecoAtendimentoDomi = 0) Then
    If CurrentQuery.FieldByName("EHTITULAR").AsString = "N" Then
      Dim qBuscaTitular As BPesquisa
      Dim qHandleEndTitular As BPesquisa
      Dim vBeneficiarioTitular As Long

      Set qBuscaTitular = NewQuery

      qBuscaTitular.Active = False
      qBuscaTitular.Clear
      qBuscaTitular.Add("SELECT B.HANDLE HBENEFICIARIOTITULAR ")
         qBuscaTitular.Add("  FROM SAM_BENEFICIARIO   B          ")
        qBuscaTitular.Add(" WHERE B.CONTRATO= :CONTRATO         ")
        qBuscaTitular.Add("   AND B.FAMILIA = :FAMILIA          ")
        qBuscaTitular.Add("   AND B.EHTITULAR= 'S'              ")
        qBuscaTitular.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
        qBuscaTitular.ParamByName("FAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
        qBuscaTitular.Active = True

        vBeneficiarioTitular = qBuscaTitular.FieldByName("HBENEFICIARIOTITULAR").AsInteger

      Set qHandleEndTitular = NewQuery

      If CurrentQuery.FieldByName("ENDERECORESIDENCIAL").AsInteger = 0 Then
        qHandleEndTitular.Active = False
        qHandleEndTitular.Clear
        qHandleEndTitular.Add("SELECT ENDERECORESIDENCIAL ")
        qHandleEndTitular.Add("  FROM SAM_BENEFICIARIO    ")
        qHandleEndTitular.Add(" WHERE HANDLE = :B_HANDLE  ")
        qHandleEndTitular.ParamByName("B_HANDLE").Value = vBeneficiarioTitular
        qHandleEndTitular.Active = True
        viHEnderecoResidencial = qHandleEndTitular.FieldByName("ENDERECORESIDENCIAL").AsInteger
      End If

      If CurrentQuery.FieldByName("ENDERECOCOMERCIAL").AsInteger = 0 Then
        qHandleEndTitular.Active = False
        qHandleEndTitular.Clear
        qHandleEndTitular.Add("SELECT ENDERECOCOMERCIAL   ")
        qHandleEndTitular.Add("  FROM SAM_BENEFICIARIO   ")
        qHandleEndTitular.Add(" WHERE HANDLE = :B_HANDLE ")
        qHandleEndTitular.ParamByName("B_HANDLE").Value = vBeneficiarioTitular
        qHandleEndTitular.Active = True
        viHEnderecoComercial = qHandleEndTitular.FieldByName("ENDERECOCOMERCIAL").AsInteger
      End If

      If CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger = 0 Then
        qHandleEndTitular.Active = False
        qHandleEndTitular.Clear
        qHandleEndTitular.Add("SELECT ENDERECOCORRESPONDENCIA  ")
        qHandleEndTitular.Add("  FROM SAM_BENEFICIARIO          ")
        qHandleEndTitular.Add(" WHERE HANDLE = :B_HANDLE        ")
        qHandleEndTitular.ParamByName("B_HANDLE").Value = vBeneficiarioTitular
        qHandleEndTitular.Active = True
        viHEnderecoCorrespondencia = qHandleEndTitular.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger
      End If

      If CurrentQuery.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger = 0 Then
        qHandleEndTitular.Active = False
        qHandleEndTitular.Clear
        qHandleEndTitular.Add("SELECT ENDERECOATENDIMENTODOMICILIAR  ")
        qHandleEndTitular.Add("  FROM SAM_BENEFICIARIO          ")
        qHandleEndTitular.Add(" WHERE HANDLE = :B_HANDLE        ")
        qHandleEndTitular.ParamByName("B_HANDLE").Value = vBeneficiarioTitular
        qHandleEndTitular.Active = True
        viHEnderecoAtendimentoDomi = qHandleEndTitular.FieldByName("ENDERECOATENDIMENTODOMICILIAR").AsInteger
      End If

      qBuscaTitular.Active = False
      Set qBuscaTitular = Nothing
      qHandleEndTitular.Active = False
      Set qHandleEndTitular = Nothing
    End If
  End If


  vdllbsben001.CarregaRotulosEndereco(CurrentSystem, _
                  viHEnderecoResidencial, _
                  "ROTULORES", _
                  vcRotulos)
  vdllbsben001.CarregaRotulosEndereco(CurrentSystem, _
                  viHEnderecoComercial, _
                  "ROTULOCOM", _
                  vcRotulos)
  vdllbsben001.CarregaRotulosEndereco(CurrentSystem, _
                  viHEnderecoCorrespondencia, _
                  "ROTULOCORRESP", _
                  vcRotulos)
  vdllbsben001.CarregaRotulosEndereco(CurrentSystem, _
                  viHEnderecoAtendimentoDomi, _
                  "ROTULOATENDDOM", _
                  vcRotulos)

  ROTULORES1.Text = vcRotulos.Field("ROTULORES1").AsString
  ROTULORES2.Text = vcRotulos.Field("ROTULORES2").AsString
  ROTULORES3.Text = vcRotulos.Field("ROTULORES3").AsString
  ROTULORES4.Text = vcRotulos.Field("ROTULORES4").AsString
  ROTULORES5.Text = vcRotulos.Field("ROTULORES5").AsString

  ROTULOCOM1.Text = vcRotulos.Field("ROTULOCOM1").AsString
  ROTULOCOM2.Text = vcRotulos.Field("ROTULOCOM2").AsString
  ROTULOCOM3.Text = vcRotulos.Field("ROTULOCOM3").AsString
  ROTULOCOM4.Text = vcRotulos.Field("ROTULOCOM4").AsString
  ROTULOCOM5.Text = vcRotulos.Field("ROTULOCOM5").AsString

  ROTULOCORRESP1.Text = vcRotulos.Field("ROTULOCORRESP1").AsString
  ROTULOCORRESP2.Text = vcRotulos.Field("ROTULOCORRESP2").AsString
  ROTULOCORRESP3.Text = vcRotulos.Field("ROTULOCORRESP3").AsString
  ROTULOCORRESP4.Text = vcRotulos.Field("ROTULOCORRESP4").AsString
  ROTULOCORRESP5.Text = vcRotulos.Field("ROTULOCORRESP5").AsString

  ROTULOATENDDOM1.Text = vcRotulos.Field("ROTULOATENDDOM1").AsString
  ROTULOATENDDOM2.Text = vcRotulos.Field("ROTULOATENDDOM2").AsString
  ROTULOATENDDOM3.Text = vcRotulos.Field("ROTULOATENDDOM3").AsString
  ROTULOATENDDOM4.Text = vcRotulos.Field("ROTULOATENDDOM4").AsString
  ROTULOATENDDOM5.Text = vcRotulos.Field("ROTULOATENDDOM5").AsString

  Set vdllbsben001 = Nothing
End Sub

Public Function ValidaNumero(pvTexto As Variant) As Boolean
  On Error GoTo erro

  CLng(pvTexto)

  ValidaNumero = True

  Exit Function

  erro:
  ValidaNumero = False
End Function

Public Sub TABLE_NewRecord()
'Isto foi feito para manter o fluxo onde ao incluir um novo registro no desktop
'o sistema abria automaticamente a interface de digitação de beneficiários
  If VisibleMode Then
    BOTAOMATRICULA_OnClick
  End If
End Sub

Public Sub TABLE_UpdateRequired()
  Dim query As BPesquisa

  Set query = NewQuery

  query.Clear
  query.Add("SELECT NOME, Z_NOME FROM SAM_MATRICULA WHERE HANDLE = :MATRICULA")
  query.ParamByName("MATRICULA").AsInteger = CurrentQuery.FieldByName("MATRICULA").AsInteger
  query.Active = True

  CurrentQuery.FieldByName("NOME").AsString = query.FieldByName("NOME").AsString
  CurrentQuery.FieldByName("Z_NOME").AsString = query.FieldByName("Z_NOME").AsString

  query.Active = False
  query.Clear
  query.Add("SELECT CONVENIO, NAOTEMCARENCIA FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
  query.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  query.Active = True

  CurrentQuery.FieldByName("CONVENIO").AsInteger = query.FieldByName("CONVENIO").AsInteger
  If query.FieldByName("NAOTEMCARENCIA").AsString = "S" Then
    CurrentQuery.FieldByName("NAOTEMCARENCIA").AsString = query.FieldByName("NAOTEMCARENCIA").AsString
  End If

  query.Active = False

  Set query = Nothing
End Sub

Public Sub TIPODEPENDENTE_OnPopup(ShowPopup As Boolean)
  If CurrentQuery.FieldByName("DATAADESAO").IsNull Then
    ShowPopup = False
    TIPODEPENDENTE.LocalWhere = "1 = 2"

    bsShowMessage("Para selecionar um tipo de dependente é obrigatório informar a data de adesão", "I")

    DATAADESAO.SetFocus
  Else
    TIPODEPENDENTE.LocalWhere = "SAM_CONTRATO_TPDEP.DATAINICIAL <= " + SQLDate(CurrentQuery.FieldByName("DATAADESAO").AsDateTime) + _
                    "AND (DATAFINAL IS NULL OR DATAFINAL >= " + SQLDate(CurrentQuery.FieldByName("DATAADESAO").AsDateTime) + " )"
  End If
End Sub

Public Sub TITULARCONJUGE_OnPopup(ShowPopup As Boolean)
  Dim vsTitulo As String
  Dim vsGrid As String
  Dim vsColunas As String
  Dim vsTabela As String
  Dim vsWhere As String
  Dim qAux As BPesquisa
  Dim viHandle As Long
  Dim dllProcura As Object

  Set qAux = NewQuery

  ShowPopup = False

  vsTitulo  = "Procura por titular cônjuge"
  vsGrid    = "Contrato|Beneficiário|Matrícula funcional|Nome|Data adesão"
  vsColunas = "C.CONTRATO|SAM_BENEFICIARIO.BENEFICIARIO|SAM_BENEFICIARIO.MATRICULAFUNCIONAL|SAM_BENEFICIARIO.NOME|SAM_BENEFICIARIO.DATAADESAO"
  vsTabela  = "SAM_BENEFICIARIO|SAM_CONTRATO C[C.HANDLE = SAM_BENEFICIARIO.CONTRATO]|SAM_MATRICULA M[M.HANDLE = SAM_BENEFICIARIO.MATRICULA]"
  vsWhere   = "SAM_BENEFICIARIO.EHTITULAR = 'S' AND SAM_BENEFICIARIO.DATACANCELAMENTO IS NULL"
  vsWhere   = vsWhere + " AND (NOT EXISTS (SELECT X.HANDLE FROM SAM_BENEFICIARIO X WHERE X.TITULARCONJUGE = SAM_BENEFICIARIO.HANDLE)"
  vsWhere   = vsWhere + "      OR SAM_BENEFICIARIO.HANDLE = " + CStr(giTitularConjuge) + ")"

  qAux.Active = False

  qAux.Clear
  qAux.Add("SELECT PERMITIRASSOCIACAOTITMESMOSEXO")
  qAux.Add("  FROM SAM_PARAMETROSBENEFICIARIO    ")

  qAux.Active = True

  If (qAux.FieldByName("PERMITIRASSOCIACAOTITMESMOSEXO").AsString = "N") Then
    qAux.Active = False
    qAux.Clear
    qAux.Add("SELECT SEXO               ")
    qAux.Add("  FROM SAM_MATRICULA      ")
    qAux.Add(" WHERE HANDLE = :MATRICULA")

    qAux.ParamByName("MATRICULA").AsInteger = giHandleMatricula

    qAux.Active = True

    If (qAux.FieldByName("SEXO").AsString = "M") Then
      vsWhere = vsWhere + " AND M.SEXO = 'F'"    ' Tabela M é a SAM_MATRICULA, que está la em cima
    Else
      vsWhere = vsWhere + " AND M.SEXO = 'M'"
    End If

    Set dllProcura = CreateBennerObject("PROCURA.Procurar")

    viHandle = dllProcura.Exec(CurrentSystem, _
                   vsTabela, _
                   vsColunas, _
                   3, _
                   vsGrid, _
                   vsWhere, _
                   vsTitulo, _
                   False, "")

    Set dllProcura = Nothing

    If (viHandle <> 0) Then
      If (CurrentQuery.State = 3) Or _
         (CurrentQuery.State = 2) Then
        CurrentQuery.FieldByName("TITULARCONJUGE").AsInteger = viHandle
      Else
        bsShowMessage("O registro não está em edição!", "I")
      End If
    End If
  End If
End Sub

Public Sub PreparaNumeracaoBenef
  If CurrentQuery.State = 1 Then
    CODIGODEPENDENTE.ReadOnly = True
  Else
  Dim query As BPesquisa

  Set query = NewQuery

  query.Clear

  query.Add("SELECT NUMEROBENEFAUTOMATICO")
  query.Add("  FROM SAM_FAMILIA        ")
  query.Add(" WHERE (HANDLE = :HANDLE)   ")

  query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("FAMILIA").AsInteger

  query.Active = True

  If (query.FieldByName("NUMEROBENEFAUTOMATICO").AsString = "N") Then
    CODIGODEPENDENTE.ReadOnly = False
  Else
    CODIGODEPENDENTE.ReadOnly = True
  End If
  End If
End Sub

Public Sub VerificaLicenca
  Dim query As BPesquisa

  Set query = NewQuery

  query.Active = False

  query.Clear
  query.Add("SELECT DATAFINAL")
  query.Add("  FROM SAM_BENEFICIARIO_LICENCA")
  query.Add(" WHERE (BENEFICIARIO = :BENEF) ")
  query.Add("   AND (DATAINICIAL <= :HOJE)  ")
  query.Add("   AND (DATAFINAL Is Null    ")
  query.Add("    OR  DATAFINAL >= :HOJE)    ")

  query.ParamByName("BENEF").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  query.ParamByName("HOJE").AsDateTime = ServerDate

  query.Active = True

  If Not query.EOF Then
    If query.FieldByName("DATAFINAL").IsNull Then
      EMLICENCA.Text = "Em licença indefinidamente"
    Else
      EMLICENCA.Text = "Em licença até " + query.FieldByName("DATAFINAL").AsString
    End If
  Else
    EMLICENCA.Text = ""
  End If
End Sub

Public Function RetornaContaFinanceira As Long
  Dim dllContaFin As Object
  Set dllContaFin = CreateBennerObject("Financeiro.Contafin")

  RetornaContaFinanceira = dllContaFin.Qual(CurrentSystem, CurrentQuery.FieldByName("FAMILIA").AsInteger, 8)

  Set dllContaFin = Nothing
End Function

Function CalculaIdadeBeneficiario(piHMatricula As Long)As Integer
  Dim vDias           As Integer
  Dim vMeses          As Integer
  Dim vAnos           As Integer
  Dim qMatricula      As Object

  Set qMatricula = NewQuery

  qMatricula.Add("SELECT DATANASCIMENTO")
  qMatricula.Add("FROM SAM_MATRICULA")
  qMatricula.Add("WHERE HANDLE = :HMATRICULA")
  qMatricula.ParamByName("HMATRICULA").AsInteger = piHMatricula
  qMatricula.Active = True

  If (qMatricula.FieldByName("DATANASCIMENTO").AsDateTime > ServerDate Or qMatricula.FieldByName("DATANASCIMENTO").IsNull) Then
    CalculaIdadeBeneficiario = 0
  Else
    DiferencaData ServerDate, qMatricula.FieldByName("DATANASCIMENTO").AsDateTime, vDias, vMeses, vAnos
    CalculaIdadeBeneficiario = vAnos
  End If

  Set qMatricula = Nothing
End Function

Public Sub DiferencaData(ByVal Data1, Data2 As Date, Dias, Meses, Anos As Integer)
  Dim DtSwap As Date
  Dim Day1, Day2, Month1, Month2, Year1, Year2 As Integer

  If Data1 >Data2 Then
    DtSwap = Data1
    Data1 = Data2
    Data2 = DtSwap
  End If

  Year1 = Val(Format(Data1, "yyyy"))
  Month1 = Val(Format(Data1, "mm"))
  Day1 = Val(Format(Data1, "dd"))

  Year2 = Val(Format(Data2, "yyyy"))
  Month2 = Val(Format(Data2, "mm"))
  Day2 = Val(Format(Data2, "dd"))

  Anos = Year2 - Year1
  Meses = 0
  Dias = 0
  If Month2 <Month1 Then
    Meses = Meses + 12
    Anos = Anos -1
  End If
  Meses = Meses + (Month2 - Month1)
  If Day2 <Day1 Then
    Dias = Dias + DiasPorMes(Year1, Val(Month1))
    If Meses = 0 Then
      Anos = Anos -1
      Meses = 11
    Else
      Meses = Meses -1
    End If
  End If
  Dias = Dias + (Day2 - Day1)
End Sub

Function DiasPorMes(ByVal Ano, Mes As Integer)As Integer
  Dim Meses31 As String
  Dim Meses30 As String

  Meses31 = "'1','3','5','7','8','10','12'"
  Meses30 = "'4','6','9','11'"

  If InStr(Meses31, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 31
  ElseIf InStr(Meses30, "'" + Str(Mes) + "'")>0 Then
    DiasPorMes = 30
  Else
    If Ano Mod 4 = 0 Then
      DiasPorMes = 29
    Else
      DiasPorMes = 28
    End If
  End If
End Function

Public Sub BOTAOMIGRAR_OnClick()
  BOTAOMIGRAR.Visible = False

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If
  If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("O Beneficiário já está cancelado !", "I")
    Exit Sub
  End If


  If Not VisibleMode Then
   Dim vsMensagem As String
   Dim viRetorno As Long
   Dim vcContainer As CSDContainer
   Dim BSINTERFACE0002 As Object

   Set vcContainer = NewContainer
   Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

   viRetorno = BSINTERFACE0002.Exec(CurrentSystem, _
								   1, _
								   "TV_TESTE_1", _
								   "Migração Beneficiário", _
								   0, _
								   120, _
								   230, _
								   False, _
								   vsMensagem, _
								   vcContainer)

   Set vcContainer = Nothing

   Select Case viRetorno
	 Case -1
		bsShowMessage("Operação cancelada pelo usuário!", "I")
     Case 1
		bsShowMessage(vsMensagem, "I")
   End Select

   Set BSINTERFACE0002 = Nothing

   CurrentQuery.Active = False
   CurrentQuery.Active = True

   Exit Sub
  End If



Dim Interface As Object
  Dim mensagem As String
  Set Interface = CreateBennerObject("CONTRATO.BENEFICIARIO")

  On Error GoTo erro
    mensagem = Interface.Migra(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    bsShowMessage(mensagem, "I")

    Set Interface = Nothing
    CurrentQuery.Active = False
    CurrentQuery.Active = True
    Exit Sub

  Erro:
    Set Interface = Nothing
    CurrentQuery.Active = False
    CurrentQuery.Active = True

    bsShowMessage(Err.Description, "I")

End Sub

Public Sub BOTAOALERTAENDCORRESP_OnClick()
  If CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").IsNull Then
    bsShowMessage("Beneficiário  " + CurrentQuery.FieldByName("NOME").AsString + " não possui endereço de correspondência.","I")
  Else
    Dim vdllbsben001 As Object
    Set vdllbsben001 = CreateBennerObject("BSBen001.BeneficiarioAux")
    vdllbsben001.AlertaEndCorresp(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger)
    Set vdllbsben001 = Nothing

    CurrentQuery.Active = False
    CurrentQuery.Active = True

  End If
End Sub

Public Sub BOTAOATUALIZAADESAO_OnClick()
  'sms 24757
  Dim vdllbsben001 As Object

  Set vdllbsben001 = CreateBennerObject("BSBen001.BeneficiarioAux")
  vdllbsben001.DataAdesao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set vdllbsben001 = Nothing
End Sub

Private Function RetornaResponsavelLegal As String
  Dim vEspecificoDLL As Object
    Set vEspecificoDLL = CreateBennerObject("Especifico.uEspecifico")
  RetornaResponsavelLegal = vEspecificoDLL.BEN_RetornaResponsavelLegal(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set vEspecificoDLL = Nothing
End Function

Public Sub ExecutaBotaoAttAdesao()
  Dim vdllbsben001 As Object
  Dim vsMensagem As String
  Dim vHbenef As Integer

  vHbenef = CurrentQuery.FieldByName("HANDLE").AsInteger

  Set vdllbsben001 = CreateBennerObject("BSBen001.BeneficiarioAux")
  vsMensagem = vdllbsben001.VerificarBeneficiarioBloqueado(CurrentSystem, vHbenef)

  If vsMensagem <> "" Then
    bsShowMessage(vsMensagem, "I")
    Err.Raise(vbsUserException, "", vsMensagem)
  End If

  vdllbsben001.DataAdesaoWeb(CurrentSystem, CurrentVirtualQuery.FieldByName("NOVADATAADESAO").AsDateTime, vHbenef)
  Set vdllbsben001 = Nothing

End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)

  Select Case CommandID
	Case "BOTAOPROPAGARENDERECO"
		BOTAOPROPAGARENDERECO_OnClick
	Case "BOTAOMIGRAR"
		BOTAOMIGRAR_OnClick
	Case "BOTAOVERIFICARINADIMPLENCIA"
        BOTAOVERIFICARINADIMPLENCIA_OnClick
	Case "BOTAOATUALIZARADESAO"
		ExecutaBotaoAttAdesao

 End Select
End Sub
