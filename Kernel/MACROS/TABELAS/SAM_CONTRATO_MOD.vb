'HASH: F7F0863D3715BF5F994511CA1CFEE37C
'Macro: SAM_CONTRATO_MOD
'#Uses "*UltimoDiaCompetencia"
'#Uses "*bsShowMessage"
Option Explicit
Dim vEstadoTabela As Long
Dim vPrecoPorTipoDepAnterior As String
Dim vPrecoPorValorOuCotaAnterior As String
Dim vAlteracao As Boolean
Dim vDataAdesao As Date

Dim Voperadora As Integer
Dim vCompFinal As Date


Public Sub BOTAOBENEFICIARIOSATIVOS_OnClick()
  'Daniela Zardo -18/07/2002
  Dim qModulo As Object
  Set qModulo = NewQuery
  Dim vGrupoFamiliar As Double

  'Número de Beneficiários desse contrato com o módulo ativo
  qModulo.Clear
  qModulo.Add("SELECT COUNT(*) MODULO FROM SAM_BENEFICIARIO_MOD WHERE MODULO = :HCONTRATOMOD And CONTRATO = :HCONTRATO And DATACANCELAMENTO Is Null")

  qModulo.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qModulo.ParamByName("HCONTRATOMOD").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

  qModulo.Active = True
  bsShowMessage("Número de Beneficiáros com o módulo ativo:" + Str(qModulo.FieldByName("MODULO").AsInteger), "I")


  qModulo.Active = False
  Set qModulo = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()


	If VisibleMode Then

		If CurrentQuery.State = 1 Then
    		Dim Interface As Object
	    	If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull  Then

  			    Dim vsMensagemErro As String
    			Dim viRetorno As Integer
    			Dim vvContainer As CSDContainer

		    	Set vvContainer = NewContainer

    			Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


			   	viRetorno = Interface.Exec(CurrentSystem, _
    									1, _
                                    	"TV_FORM0015", _
           	                        	"Cancelamento do Módulo do Contrato", _
               	                    	0, _
                   	                	200, _
                       	            	450, _
                           	        	False, _
                               	    	vsMensagemErro, _
                                   		vvContainer)

   				Select Case viRetorno
      				Case -1
   						bsShowMessage("Operação cancelada pelo usuário!", "I")
  					Case  0
   	  					'bsShowMessage("Opção selecionada" + vvContainer.Field("OPCAO").AsString , "I")
  					Case  1
   	  					bsShowMessage(vsMensagemErro, "I")
				End Select


				Set Interface = Nothing
				CurrentQuery.Active = False
    	 		CurrentQuery.Active = True

    		Else
    			bsShowMessage("Contrato já cancelado!", "I")
    		End If

  		End If

	End If
End Sub

Public Sub BOTAOPROPAGAR_OnClick()
  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                   0, _
                                   0, _
                                   RecordHandleOfTable("SAM_CONTRATO"), _
                                   vDataFinalSuspensao)Then
     bsShowMessage("Não é permitido propagar o módulo por motivo de suspensão!", "I")
     Exit Sub
   End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------

  'Daniela Zardo -02/07/2002
    If CurrentQuery.State = 1 Then
    Dim Interface As Object

    Set Interface = CreateBennerObject("CONTRATO.Propagar")'Anderson sms 21638(Plano)
    Interface.Inclui(CurrentSystem, CurrentQuery.FieldByName("PLANO").AsInteger, CurrentQuery.FieldByName("CONTRATO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing

    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If

End Sub

Public Sub BOTAOREATIVAR_OnClick()

If VisibleMode Then

  	If CurrentQuery.State = 1 Then
	    Dim Interface As Object
	    If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    	  	Dim SQL As Object
	      	Set SQL = NewQuery
      		SQL.Add("SELECT TABTIPOCONTRATO, DATACANCELAMENTO FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO")
	      	SQL.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
      		SQL.Active = True
      		If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
        		bsShowMessage("Não é permitido reativar módulos nesse contrato - Contrato está Cancelado!", "I")
        		Exit Sub
      		End If
			    Dim vsMensagemErro As String
    			Dim viRetorno As Integer
    			Dim vvContainer As CSDContainer

		    	Set vvContainer = NewContainer

    			Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


			   	viRetorno = Interface.Exec(CurrentSystem, _
    									1, _
                                    	"TV_FORM0010", _
           	                        	"Reativação do Módulo do Contrato", _
               	                    	0, _
                   	                	120, _
                       	            	280, _
                           	        	False, _
                               	    	vsMensagemErro, _
                                   		vvContainer)

   				Select Case viRetorno
      				Case -1
   						bsShowMessage("Operação cancelada pelo usuário!", "I")
  					Case  0
   	  					'bsShowMessage("Opção selecionada" + vvContainer.Field("OPCAO").AsString , "I")
  					Case  1
   	  					bsShowMessage(vsMensagemErro, "I")
				End Select

    	  Set SQL = Nothing
    	  Set Interface = Nothing
    	  CurrentQuery.Active = False
    	  CurrentQuery.Active = True
 	   Else
 	   	  bsShowMessage("Módulo do Contrato não Cancelado!","I")
 	   End If

  	End If

 End If
End Sub


Public Sub BOTAOTRANSFEREMODULO_OnClick()
  'Valeska - sms 16738
  Dim vDataFinalSuspensao As Date
  Dim SQL As Object
  Set SQL = NewQuery
  Dim SQL1 As Object
  Set SQL1 = NewQuery

  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, 0, 0, RecordHandleOfTable("SAM_CONTRATO"), vDataFinalSuspensao) Then
    bsShowMessage("Não é permitido transferir o módulo por motivo de suspensão!", "I")
    Exit Sub
  End If
  Set BSBen001Dll = Nothing

  If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("Não é possível fazer esta transferência. Este contrato está cancelado!", "E")
    Exit Sub
  End If




  If CurrentQuery.State = 1 Then
    Dim QUERY As Object
    Set QUERY = NewQuery
    Dim vTransferirANS As String

    QUERY.Active = False
    QUERY.Clear
    QUERY.Add("SELECT NAOTRANSFERIRMODULOSEXPORTADOS")
    QUERY.Add("  FROM SAM_PARAMETROSBENEFICIARIO")
    QUERY.Active = True

    vTransferirANS = QUERY.FieldByName("NAOTRANSFERIRMODULOSEXPORTADOS").AsString

    If vTransferirANS = "S" Then

      'Se o beneficiário For titular e o parâmetro "Cancelar módulos do mesmo registro" estiver
      'marcado em SAM_REGISTROMS, todos os módulos dos dependentes Do titular serão cancelados

      QUERY.Active = False
      QUERY.Clear
      QUERY.Add("SELECT SR.EXPORTARBENEFICIARIOS ")
      QUERY.Add("  FROM SAM_CONTRATO C,")
      QUERY.Add("       SAM_REGISTROMS SR,")
      QUERY.Add("       SAM_CONTRATO_MOD SCM")
      QUERY.Add(" WHERE C.HANDLE = SCM.CONTRATO")
      QUERY.Add("   AND C.CONVENIO = SR.CONVENIO")
      QUERY.Add("   AND SR.HANDLE = SCM.REGISTROMS")
      QUERY.Add("   AND SCM.HANDLE = :MODULO")
      QUERY.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsString
      QUERY.Active = True
      If QUERY.FieldByName("EXPORTARBENEFICIARIOS").AsString = "S" Then
        bsShowMessage("Módulos que estão configurados para serem exportados para a ANS não podem sofrer transferência." + Chr(13) + _
               " O beneficiário deverá ser migrado !", "I")
        QUERY.Active = False
        Exit Sub
      End If
      QUERY.Active = False
    End If


  	Dim INTERFACE0002 As Object
  	Dim vsMensagem As String
  	Dim vcContainer As CSDContainer


  	Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")
  	Set vcContainer = NewContainer

  	INTERFACE0002.Exec(CurrentSystem, _
					   	1, _
					   	"TV_FORM0087", _
					   	"Tranferência de Módulos",  _
					   	CurrentQuery.FieldByName("HANDLE").AsInteger, _
					   	400, _
					   	350, _
					   	False, _
					   	vsMensagem, _
					   	vcContainer)

  	Set INTERFACE0002 = Nothing
  End If

End Sub


Public Sub MODULO_OnChange()
  Dim SQL As Object

  Set SQL = NewQuery

  SQL.Add("SELECT P.REGISTROMS")
  SQL.Add("FROM SAM_CONTRATO C, SAM_PLANO_MOD P")
  SQL.Add("WHERE C.HANDLE = :HCONTRATO")
  SQL.Add("  AND P.PLANO = C.PLANO")
  SQL.Add("  AND P.MODULO = :HMODULO")
  SQL.ParamByName("HCONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
  SQL.ParamByName("HMODULO").Value = CurrentQuery.FieldByName("MODULO").AsInteger
  SQL.Active = True

  If Not SQL.FieldByName("REGISTROMS").IsNull Then
    CurrentQuery.FieldByName("REGISTROMS").AsInteger = SQL.FieldByName("REGISTROMS").AsInteger
  Else
    CurrentQuery.FieldByName("REGISTROMS").Clear
  End If

  Set SQL = Nothing
End Sub



Public Sub PLANOFRANQUIA_OnChange()
  CurrentQuery.FieldByName("GRUPOPFCALCULOFRQURGEMERG").Clear
End Sub




Public Sub REGISTROMS_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT CONVENIO")
  SQL.Add("FROM SAM_CONTRATO")
  SQL.Add("WHERE HANDLE = :HCONTRATO")
  SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  SQL.Active = True

  'Daniela -SMS 12220 -Convênio no registro da ANS
  vCriterio = "CONVENIO = " + SQL.FieldByName("CONVENIO").AsString

  Set SQL = Nothing

  vColunas = "SAM_REGISTROMS.REGISTROMS|SAM_REGISTROMS.DESCRICAO|SAM_SEGMENTACAO.DESCRICAO|SAM_REGISTROMS.DATAVENCIMENTO"

  vCampos = "Registro|Descrição|Segmentação|Vencimento"

  vHandle = Interface.Exec(CurrentSystem, "SAM_REGISTROMS|SAM_SEGMENTACAO[SAM_REGISTROMS.SEGMENTACAO = SAM_SEGMENTACAO.HANDLE]", vColunas, 1, vCampos, vCriterio, "Registro no Ministério da Saúde", True, "")

  If vHandle <>0 Then
    '    CurrentQuery.Edit
    CurrentQuery.FieldByName("REGISTROMS").Value = vHandle
  End If

  Set Interface = Nothing
End Sub

Public Sub TABLE_AfterEdit()
  vDataAdesao = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
End Sub

Public Sub TABLE_AfterPost()
  Dim vPrimeiraCompetencia As Date
  Dim vUltimaCompetencia As Date
  Dim SQL As Object

  If vAlteracao Then
    If CurrentQuery.FieldByName("PRECOPORTIPODEPENDENTE").AsString <>vPrecoPorTipoDepAnterior Then
      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
      SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM,")
      SQL.Add("     SAM_CONTRATO_MOD CM")
      SQL.Add("WHERE CM.HANDLE = :HCONTRATOMOD")
      SQL.Add("  AND BM.MODULO = CM.HANDLE")
      SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
      SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
      SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
      SQL.ParamByName("HCONTRATOMOD").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.ParamByName("HOJE").Value = ServerDate
      SQL.Active = True

      vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
      vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

      If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then

        SQL.Clear
        SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
        SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
        SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATOINICIAL, CONTRATOFINAL,")
        SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
        SQL.Add("VALUES")
        SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 2,")
        SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HCONTRATO,")
        SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

        SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
        SQL.ParamByName("DATAROTINA").Value = ServerDate
        SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
        SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
        SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
        SQL.ParamByName("USUARIO").Value = CurrentUser
        SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
        SQL.ParamByName("DESCRICAO").Value = "Alteração no parâmetro 'Preço por Tipo Dependente'"

        SQL.ExecSQL

      End If

      Set SQL = Nothing
    ElseIf CurrentQuery.FieldByName("PRECOPORVALOROUCOTA").AsString <>vPrecoPorValorOuCotaAnterior Then
      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT MIN(COMPETENCIA) PRIMEIRACOMPETENCIA, MAX(COMPETENCIA) ULTIMACOMPETENCIA")
      SQL.Add("FROM SFN_FATURA_LANC_MOD A, SAM_BENEFICIARIO_MOD BM,")
      SQL.Add("     SAM_CONTRATO_MOD CM")
      SQL.Add("WHERE CM.HANDLE = :HCONTRATOMOD")
      SQL.Add("  AND BM.MODULO = CM.HANDLE")
      SQL.Add("  AND (   BM.DATACANCELAMENTO IS NULL")
      SQL.Add("       OR BM.DATACANCELAMENTO >= :HOJE)")
      SQL.Add("  AND A.BENEFICIARIOMOD = BM.HANDLE")
      SQL.ParamByName("HCONTRATOMOD").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.ParamByName("HOJE").Value = ServerDate
      SQL.Active = True

      vPrimeiraCompetencia = SQL.FieldByName("PRIMEIRACOMPETENCIA").AsDateTime
      vUltimaCompetencia = SQL.FieldByName("ULTIMACOMPETENCIA").AsDateTime

      If Not SQL.FieldByName("PRIMEIRACOMPETENCIA").IsNull Then

        SQL.Clear
        SQL.Add("INSERT INTO SAM_ROTINARECALCULOMENSALID")
        SQL.Add("(HANDLE, CODIGO, DESCRICAO, DATAROTINA, TABRECALCULAR,")
        SQL.Add(" COMPETENCIAINICIAL, COMPETENCIAFINAL, CONTRATOINICIAL, CONTRATOFINAL,")
        SQL.Add(" USUARIO, DATAINCLUSAO, SITUACAOPROCESSAMENTO, SITUACAOFATURAMENTO)")
        SQL.Add("VALUES")
        SQL.Add("(:HANDLE, :HANDLE, :DESCRICAO, :DATAROTINA, 2,")
        SQL.Add(" :COMPETENCIAINICIAL, :COMPETENCIAFINAL, :HCONTRATO, :HCONTRATO,")
        SQL.Add(" :USUARIO, :DATAINCLUSAO, '1', '1')")

        SQL.ParamByName("HANDLE").Value = NewHandle("SAM_ROTINARECALCULOMENSALID")
        SQL.ParamByName("DATAROTINA").Value = ServerDate
        SQL.ParamByName("COMPETENCIAINICIAL").Value = vPrimeiraCompetencia
        SQL.ParamByName("COMPETENCIAFINAL").Value = vUltimaCompetencia
        SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
        SQL.ParamByName("USUARIO").Value = CurrentUser
        SQL.ParamByName("DATAINCLUSAO").Value = ServerDate
        SQL.ParamByName("DESCRICAO").Value = "Alteração no parâmetro 'Preço por Valor ou Cota'"

        SQL.ExecSQL

      End If

      Set SQL = Nothing
    End If
  End If

  If vEstadoTabela = 3 Then
    Dim Interface As Object
    Set Interface = CreateBennerObject("CONTRATO.ContratoModulo")
    Interface.Inclui(CurrentSystem, CurrentQuery.FieldByName("CONTRATO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, CurrentQuery.FieldByName("MODULO").AsInteger)
    Set Interface = Nothing
  End If

  'Anderson sms 21638
  If vEstadoTabela = 3 Then
    Dim vFilial As Integer
    Dim SQL2 As Object
    Set SQL2 = NewQuery

    'Verifica se o centro de custo do contrato é por módulo
    SQL2.Clear
    SQL2.Add("SELECT TABCENTROCUSTO, TIPOCONTRATACAOEMPRESARIAL FROM SAM_CONTRATO ")
    SQL2.Add("WHERE HANDLE = :CONTRATO ")
    SQL2.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
    SQL2.Active = True

    If SQL2.FieldByName("TABCENTROCUSTO").AsInteger = 4 Then
      Dim qPlanoCentroCusto As Object
      Dim qInsCentroCustoMod As Object
      Set qInsCentroCustoMod = NewQuery
      Set qPlanoCentroCusto = NewQuery

      qPlanoCentroCusto.Clear
      qPlanoCentroCusto.Add("SELECT  PC.FILIAL, PC.CENTROCUSTO CENTROCUSTOFILIAL, PCM.CENTROCUSTO CENTROCUSTOMOD ")
      qPlanoCentroCusto.Add("FROM SAM_PLANO_CENTROCUSTO PC, SAM_PLANO_CENTROCUSTO_MODULO PCM ")
      qPlanoCentroCusto.Add("WHERE PC.PLANO = :PLANO And PC.HANDLE = PCM.PLANOCENTROCUSTO ")
      qPlanoCentroCusto.Add("AND PCM.MODULO = :MODULO AND PC.TIPOCONTRATACAOEMPRESARIAL = :TIPOCONTRATACAO ")
      qPlanoCentroCusto.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
      qPlanoCentroCusto.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
      qPlanoCentroCusto.ParamByName("TIPOCONTRATACAO").AsString = SQL2.FieldByName("TIPOCONTRATACAOEMPRESARIAL").AsString
      qPlanoCentroCusto.Active = True
      vFilial = qPlanoCentroCusto.FieldByName("FILIAL").AsInteger

      If Not qPlanoCentroCusto.FieldByName("FILIAL").IsNull Then
        Dim qVerificaFilial As Object
        Set qVerificaFilial = NewQuery
        'Verifica se a filial do centro de custo do contrato é a mesma do plano
        qVerificaFilial.Clear
        qVerificaFilial.Add("SELECT * FROM SAM_PLANO_CENTROCUSTO ")
        qVerificaFilial.Add("WHERE PLANO = :PLANO AND FILIAL IN  ")
        qVerificaFilial.Add("(SELECT FILIAL FROM SAM_CONTRATO_CENTROCUSTO WHERE CONTRATO = :CONTRATO AND FILIAL = :FILIAL) ")
        qVerificaFilial.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
        qVerificaFilial.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
        qVerificaFilial.ParamByName("FILIAL").AsInteger = vFilial
        qVerificaFilial.Active = True
        If Not qVerificaFilial.FieldByName("HANDLE").IsNull Then
          Dim qContratoCentroCusto As Object
          Set qContratoCentroCusto = NewQuery

          'Posiciona no módulo da filial
          qPlanoCentroCusto.Clear
          qPlanoCentroCusto.Add("SELECT  PC.FILIAL, PC.CENTROCUSTO CENTROCUSTOFILIAL, PCM.CENTROCUSTO CENTROCUSTOMOD ")
          qPlanoCentroCusto.Add("FROM SAM_PLANO_CENTROCUSTO PC, SAM_PLANO_CENTROCUSTO_MODULO PCM ")
          qPlanoCentroCusto.Add("WHERE PC.PLANO = :PLANO And PC.HANDLE = PCM.PLANOCENTROCUSTO ")
          qPlanoCentroCusto.Add("AND PCM.MODULO = :MODULO AND PC.TIPOCONTRATACAOEMPRESARIAL = :TIPOCONTRATACAO ")
          qPlanoCentroCusto.Add("AND PC.FILIAL = :FILIAL ")
          qPlanoCentroCusto.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
          qPlanoCentroCusto.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
          qPlanoCentroCusto.ParamByName("TIPOCONTRATACAO").AsString = SQL2.FieldByName("TIPOCONTRATACAOEMPRESARIAL").AsString
          qPlanoCentroCusto.ParamByName("FILIAL").AsInteger = vFilial
          qPlanoCentroCusto.Active = True

          qContratoCentroCusto.Clear
          qContratoCentroCusto.Add("SELECT * ")
          qContratoCentroCusto.Add("FROM SAM_CONTRATO_CENTROCUSTO ")
          qContratoCentroCusto.Add("WHERE CONTRATO = :CONTRATO AND FILIAL = :FILIAL ")
          qContratoCentroCusto.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
          qContratoCentroCusto.ParamByName("FILIAL").AsInteger = vFilial
          qContratoCentroCusto.Active = True

          'Insere o módulo na SAM_CONTRATO_CENTROCUSTO_MOD
          qInsCentroCustoMod.Clear
          qInsCentroCustoMod.Add("INSERT INTO SAM_CONTRATO_CENTROCUSTO_MOD ")
          qInsCentroCustoMod.Add("(HANDLE, CONTRATOCENTROCUSTO, CONTRATOMODULO, CENTROCUSTO) ")
          qInsCentroCustoMod.Add("VALUES (:HANDLE, :CONTRATOCENTROCUSTO, :CONTRATOMODULO, :CENTROCUSTO) ")
          qInsCentroCustoMod.ParamByName("HANDLE").AsInteger = NewHandle("SAM_CONTRATO_CENTROCUSTO_MOD ")
          qInsCentroCustoMod.ParamByName("CONTRATOCENTROCUSTO").AsInteger = qContratoCentroCusto.FieldByName("HANDLE").AsInteger
          qInsCentroCustoMod.ParamByName("CONTRATOMODULO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          qInsCentroCustoMod.ParamByName("CENTROCUSTO").AsInteger = qPlanoCentroCusto.FieldByName("CENTROCUSTOMOD").AsInteger
          qInsCentroCustoMod.ExecSQL

        Else
          Dim nHandle As Integer
          Dim qInsCentroCustoContrato As Object
          Set qInsCentroCustoContrato = NewQuery
          nHandle = NewHandle("SAM_CONTRATO_CENTROCUSTO")
          'Insere SAM_CONTRATO_CENTROCUSTO com a filial do plano
          qInsCentroCustoContrato.Clear
          qInsCentroCustoContrato.Add("INSERT INTO SAM_CONTRATO_CENTROCUSTO ")
          qInsCentroCustoContrato.Add("(HANDLE, CONTRATO, FILIAL, CENTROCUSTO) ")
          qInsCentroCustoContrato.Add("VALUES (:HANDLE, :CONTRATO, :FILIAL, :CENTROCUSTO) ")
          qInsCentroCustoContrato.ParamByName("HANDLE").AsInteger = nHandle
          qInsCentroCustoContrato.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
          qInsCentroCustoContrato.ParamByName("FILIAL").AsInteger = qPlanoCentroCusto.FieldByName("FILIAL").AsInteger
          qInsCentroCustoContrato.ParamByName("CENTROCUSTO").AsInteger = qPlanoCentroCusto.FieldByName("CENTROCUSTOFILIAL").AsInteger
          qInsCentroCustoContrato.ExecSQL

          'Posiciona no módulo da filial
          qPlanoCentroCusto.Clear
          qPlanoCentroCusto.Add("SELECT  PC.FILIAL, PC.CENTROCUSTO CENTROCUSTOFILIAL, PCM.CENTROCUSTO CENTROCUSTOMOD ")
          qPlanoCentroCusto.Add("FROM SAM_PLANO_CENTROCUSTO PC, SAM_PLANO_CENTROCUSTO_MODULO PCM ")
          qPlanoCentroCusto.Add("WHERE PC.PLANO = :PLANO And PC.HANDLE = PCM.PLANOCENTROCUSTO ")
          qPlanoCentroCusto.Add("AND PCM.MODULO = :MODULO AND PC.TIPOCONTRATACAOEMPRESARIAL = :TIPOCONTRATACAO ")
          qPlanoCentroCusto.Add("AND PC.FILIAL = :FILIAL ")
          qPlanoCentroCusto.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
          qPlanoCentroCusto.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
          qPlanoCentroCusto.ParamByName("TIPOCONTRATACAO").AsString = SQL2.FieldByName("TIPOCONTRATACAOEMPRESARIAL").AsString
          qPlanoCentroCusto.ParamByName("FILIAL").AsInteger = vFilial
          qPlanoCentroCusto.Active = True

          'Insere o módulo na SAM_CONTRATO_CENTROCUSTO_MOD
          qInsCentroCustoMod.Clear
          qInsCentroCustoMod.Add("INSERT INTO SAM_CONTRATO_CENTROCUSTO_MOD ")
          qInsCentroCustoMod.Add("(HANDLE, CONTRATOCENTROCUSTO, CONTRATOMODULO, CENTROCUSTO) ")
          qInsCentroCustoMod.Add("VALUES (:HANDLE, :CONTRATOCENTROCUSTO, :CONTRATOMODULO, :CENTROCUSTO) ")
          qInsCentroCustoMod.ParamByName("HANDLE").AsInteger = NewHandle("SAM_CONTRATO_CENTROCUSTO_MOD ")
          qInsCentroCustoMod.ParamByName("CONTRATOCENTROCUSTO").AsInteger = nHandle
          qInsCentroCustoMod.ParamByName("CONTRATOMODULO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          qInsCentroCustoMod.ParamByName("CENTROCUSTO").AsInteger = qPlanoCentroCusto.FieldByName("CENTROCUSTOMOD").AsInteger
          qInsCentroCustoMod.ExecSQL

          Set qContratoCentroCusto = Nothing
          Set qInsCentroCustoContrato = Nothing
        End If
        Set qVerificaFilial = Nothing
      End If

      Set qInsCentroCustoMod = Nothing
      Set qPlanoCentroCusto = Nothing
    End If
    'Verifica se existe o modulo na Limitacao do plano
    SQL2.Clear
    SQL2.Add("SELECT CL.HANDLE FROM SAM_PLANO_LIMITACAO_MOD LM, ")
    SQL2.Add("SAM_PLANO_MOD PM, SAM_PLANO_LIMITACAO PL, SAM_CONTRATO_LIMITACAO CL ")
    SQL2.Add("WHERE LM.PLANOMODULO = PM.HANDLE AND LM.PLANOLIMITACAO = PL.HANDLE ")
    SQL2.Add("AND PM.PLANO = PL.PLANO AND PM.PLANO = :PLANO AND PM.MODULO = :MODULO ")
    SQL2.Add("AND CL.PLANO = PM.PLANO AND CL.LIMITACAO = PL.LIMITACAO ")
    SQL2.Add("AND CL.CONTRATO = :CONTRATO ")
    SQL2.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
    SQL2.ParamByName("MODULO").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
    SQL2.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
    SQL2.Active = True

    If Not SQL2.FieldByName("HANDLE").IsNull Then
      Dim qInsLimitacaoMod As Object
      Set qInsLimitacaoMod = NewQuery

      'Insere o módulo na SAM_CONTRATO_LIMITACAO_MOD
      While Not SQL2.EOF
        qInsLimitacaoMod.Clear
        qInsLimitacaoMod.Add("INSERT INTO SAM_CONTRATO_LIMITACAO_MOD ")
        qInsLimitacaoMod.Add("(HANDLE, CONTRATO, CONTRATOLIMITACAO, CONTRATOMODULO) ")
        qInsLimitacaoMod.Add("VALUES (:HANDLE, :CONTRATO, :CONTRATOLIMITACAO, :CONTRATOMODULO) ")
        qInsLimitacaoMod.ParamByName("HANDLE").AsInteger = NewHandle("SAM_CONTRATO_LIMITACAO_MOD")
        qInsLimitacaoMod.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
        qInsLimitacaoMod.ParamByName("CONTRATOLIMITACAO").AsInteger = SQL2.FieldByName("HANDLE").AsInteger
        qInsLimitacaoMod.ParamByName("CONTRATOMODULO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qInsLimitacaoMod.ExecSQL
        SQL2.Next
      Wend
      Set qInsLimitacaoMod = Nothing
    End If

    Set SQL2 = Nothing
  End If

End Sub

Public Sub TABLE_AfterScroll()

  If WebMode Then
  		REGISTROMS.WebLocalWhere = "A.CONVENIO IN(SELECT CONVENIO   " + _
  								   "FROM SAM_CONTRATO			  " + _
 								   "WHERE HANDLE = @CAMPO(CONTRATO) )"
  End If


  SessionVar("HPLANO") = CurrentQuery.FieldByName("PLANO").AsString
  SessionVar("HCONTRATO") = CurrentQuery.FieldByName("CONTRATO").AsString
  SessionVar("HMODULO") = CurrentQuery.FieldByName("HANDLE").AsString

  Dim SQL As Object
  Set sql = NewQuery

  Dim SQL1 As Object
  Set SQL1 = NewQuery

	SQL.Active = False
  	SQL.Clear
  	SQL.Add("SELECT MODULO, PLANO FROM SAM_CONTRATO_MOD WHERE HANDLE = :PHANDLE")
  	SQL.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  	SQL.Active = True

  	SQL1.Active = False
  	SQL1.Clear
  	SQL1.Add("SELECT CONTRATO, CONTRATANTE FROM SAM_CONTRATO WHERE HANDLE = :PHANDLE")
  	SQL1.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  	SQL1.Active = True


  	SessionVar("contrato") = SQL1.FieldByName("CONTRATO").AsString
  	SessionVar("contratante") = SQL1.FieldByName("CONTRATANTE").AsString
  	SessionVar("plano") = SQL.FieldByName("plano").AsString
  	SessionVar("modulo") = SQL.FieldByName("modulo").AsString


Set SQL = Nothing
Set SQL1= Nothing

  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    AGENTEAGENCIAVENDAS.ReadOnly = False
    TIPOCOMISSAO.ReadOnly = False
    PARCELADIAS.ReadOnly = False
    PRIMEIRAPARCELA.ReadOnly = False
    REDEDIFERENCIADA.ReadOnly = False
    REDERESTRITA.ReadOnly = False
    DATAADESAO.ReadOnly = False
  Else
    AGENTEAGENCIAVENDAS.ReadOnly = True
    TIPOCOMISSAO.ReadOnly = True
    PARCELADIAS.ReadOnly = True
    PRIMEIRAPARCELA.ReadOnly = True
    REDEDIFERENCIADA.ReadOnly = True
    REDERESTRITA.ReadOnly = True
    DATAADESAO.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    0, _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido excluir o módulo por motivo de suspensão!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)


  If WebMode Then
  	PLANOFRANQUIA.WebLocalWhere = "HANDLE IN (                                       " + _
    	                         "           SELECT PLANO                           " + _
        	                     "             FROM SAM_CONTRATO_PLANO              " + _
            	                 "            WHERE CONTRATO = @CAMPO(CONTRATO)" + _
                	             "          )

    PLANO.WebLocalWhere = "HANDLE IN (                                       " + _
    	                 "           SELECT PLANO                           " + _
        	             "             FROM SAM_CONTRATO_PLANO              " + _
            	         "            WHERE CONTRATO = @CAMPO(CONTRATO)" + _
                	     "          )

  	MODULO.WebLocalWhere = "HANDLE IN ( SELECT PM.MODULO                 " + _
    	                  "              FROM SAM_CONTRATO_PLANO CP,    " + _
        	              "                   SAM_PLANO_MOD PM          " + _
            	          "             WHERE CP.PLANO = PM.PLANO        " + _
                	      "               AND CP.CONTRATO = @CAMPO(CONTRATO)" + _
                    	  "               AND CP.PLANO    = @CAMPO(PLANO)" + _
                      	  "          )   "
    GRUPOPFCALCULOFRQURGEMERG.WebLocalWhere = "HANDLE IN ( SELECT TABELAPFEVENTO                 " + _
                                         "              FROM SAM_CONTRATO_PFEVENTO           " + _
                                         "             WHERE CONTRATO = @CAMPO(CONTRATO)" + _
                                         "               AND PLANO    = @CAMPO(PLANOFRANQUIA)" + _
                                         "          )   "

  ElseIf VisibleMode Then
  	PLANOFRANQUIA.LocalWhere = "HANDLE IN (                                       " + _
    	                         "           SELECT PLANO                           " + _
        	                     "             FROM SAM_CONTRATO_PLANO              " + _
            	                 "            WHERE CONTRATO = @CONTRATO" + _
                	             "          )

   	PLANO.LocalWhere = "HANDLE IN (                                       " + _
    	                 "           SELECT PLANO                           " + _
        	             "             FROM SAM_CONTRATO_PLANO              " + _
            	         "            WHERE CONTRATO = @CONTRATO" + _
                	     "          )

    MODULO.LocalWhere = "HANDLE IN ( SELECT B.MODULO                 " + _
    	                  "              FROM SAM_CONTRATO_PLANO A,    " + _
        	              "                   SAM_PLANO_MOD B          " + _
            	          "             WHERE A.PLANO = B.PLANO        " + _
                	      "               AND A.CONTRATO = @CONTRATO" + _
                    	  "               AND A.PLANO    = @PLANO" + _
                      	  "          )   "

    GRUPOPFCALCULOFRQURGEMERG.LocalWhere = "HANDLE IN ( SELECT TABELAPFEVENTO                 " + _
    	                                     "              FROM SAM_CONTRATO_PFEVENTO           " + _
        	                                 "             WHERE CONTRATO = @CONTRATO" + _
            	                             "               AND PLANO    = @PLANOFRANQUIA "+ _
                	                         "          )   "

  End If



  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    0, _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido editar o módulo por motivo de suspensão!", "E")
    CanContinue = False
    CurrentQuery.Cancel
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------

  vAlteracao = True
  vPrecoPorTipoDepAnterior = CurrentQuery.FieldByName("PRECOPORTIPODEPENDENTE").AsString
  vPrecoPorValorOuCotaAnterior = CurrentQuery.FieldByName("PRECOPORVALOROUCOTA").AsString
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  If WebMode Then
  	PLANOFRANQUIA.WebLocalWhere = "HANDLE IN (                                       " + _
    	                         "           SELECT PLANO                           " + _
        	                     "             FROM SAM_CONTRATO_PLANO              " + _
            	                 "            WHERE CONTRATO = @CAMPO(CONTRATO)" + _
                	             "          )

    PLANO.WebLocalWhere = "HANDLE IN (                                       " + _
    	                 "           SELECT PLANO                           " + _
        	             "             FROM SAM_CONTRATO_PLANO              " + _
            	         "            WHERE CONTRATO = @CAMPO(CONTRATO)" + _
                	     "          )

  	MODULO.WebLocalWhere = "HANDLE IN ( SELECT PM.MODULO                 " + _
    	                  "              FROM SAM_CONTRATO_PLANO CP,    " + _
        	              "                   SAM_PLANO_MOD PM          " + _
            	          "             WHERE CP.PLANO = PM.PLANO        " + _
                	      "               AND CP.CONTRATO = @CAMPO(CONTRATO)" + _
                    	  "               AND CP.PLANO    = @CAMPO(PLANO)" + _
                      	  "          )   "
    GRUPOPFCALCULOFRQURGEMERG.WebLocalWhere = "HANDLE IN ( SELECT TABELAPFEVENTO                 " + _
                                         "              FROM SAM_CONTRATO_PFEVENTO           " + _
                                         "             WHERE CONTRATO = @CAMPO(CONTRATO)" + _
                                         "               AND PLANO    = @CAMPO(PLANOFRANQUIA)" + _
                                         "          )   "

  ElseIf VisibleMode Then
  	PLANOFRANQUIA.LocalWhere = "HANDLE IN (                                       " + _
    	                         "           SELECT PLANO                           " + _
        	                     "             FROM SAM_CONTRATO_PLANO              " + _
            	                 "            WHERE CONTRATO = @CONTRATO" + _
                	             "          )

   	PLANO.LocalWhere = "HANDLE IN (                                       " + _
    	                 "           SELECT PLANO                           " + _
        	             "             FROM SAM_CONTRATO_PLANO              " + _
            	         "            WHERE CONTRATO = @CONTRATO" + _
                	     "          )

    MODULO.LocalWhere = "HANDLE IN ( SELECT B.MODULO                 " + _
    	                  "              FROM SAM_CONTRATO_PLANO C,    " + _
        	              "                   SAM_PLANO_MOD B          " + _
            	          "             WHERE C.PLANO = B.PLANO        " + _
                	      "               AND C.CONTRATO = @CONTRATO" + _
                    	  "               AND C.PLANO    = @PLANO" + _
                      	  "          )   "

    GRUPOPFCALCULOFRQURGEMERG.LocalWhere = "HANDLE IN ( SELECT TABELAPFEVENTO                 " + _
    	                                     "              FROM SAM_CONTRATO_PFEVENTO           " + _
        	                                 "             WHERE CONTRATO = @CONTRATO" + _
            	                             "               AND PLANO    = @PLANOFRANQUIA "+ _
                	                         "          )   "

  End If






  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    0, _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido inserir o módulo por motivo de suspensão!", "E")
    CanContinue = False
    CurrentQuery.Cancel
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------

  vAlteracao = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim SqlOperadora As BPesquisa
Dim Sqltipo As BPesquisa

  'Daniela Zardo -03/07/2002
  If(CurrentQuery.FieldByName("ACEITATITULAR").AsString = "N") _
     And(CurrentQuery.FieldByName("ACEITADEPENDENTES").AsString = "N") _
     And(CurrentQuery.FieldByName("ACEITAAGREGADOS").AsString = "N")Then
  bsShowMessage("Um dos campos Aceita Titular, Dependente ou Agregado devem estar ativos!", "E")
  CanContinue = False
End If

'Anderson 04/08/03 sms 17035
'---------------------------------------------------------------------------------------------------------------------------
Set SqlOperadora = NewQuery
Set Sqltipo = NewQuery

If CurrentQuery.State = 3 Then
  Sqltipo.Active = False
  Sqltipo.Clear
  Sqltipo.Add("SELECT NAOREGISTRARNOMS FROM SAM_CONTRATO WHERE HANDLE = :CONTRATO ")
  Sqltipo.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  Sqltipo.Active = True


  If Sqltipo.FieldByName("NAOREGISTRARNOMS").AsString = "N" Then

    Sqltipo.Active = False
    Sqltipo.Clear
    Sqltipo.Add("SELECT C.OPERADORA, SO.TABTIPO, SO.HANDLE, SO.ADMINISTRADORA,SC.CONVENIO ")
    Sqltipo.Add("  FROM SAM_CONVENIO C  ,                                                 ")
    Sqltipo.Add("       SAM_CONTRATO SC ,                                                 ")
    Sqltipo.Add("       SAM_OPERADORA SO                                                  ")
    Sqltipo.Add(" WHERE SC.CONVENIO = C.HANDLE                                            ")
    Sqltipo.Add("   And C.OPERADORA = SO.HANDLE                                           ")
    Sqltipo.Add("   AND SC.HANDLE = :CONTRATO                                             ")
    Sqltipo.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
    Sqltipo.Active = True

    If Sqltipo.FieldByName("tabtipo").Value = 1 Then
      Voperadora = Sqltipo.FieldByName("OPERADORA").AsInteger

      SqlOperadora.Active = False
      SqlOperadora.Clear
      SqlOperadora.Add("SELECT MAX(DATAFINAL)DATAFINAL     ")
      SqlOperadora.Add("  FROM SAM_MSPROCESSO              ")
      SqlOperadora.Add(" WHERE OPERADORA = :OPERADORA      ")
      SqlOperadora.Add("   AND SITUACAO  = 'P'             ")
      SqlOperadora.Add("   AND TABTIPOOPERADORA = :TABTIPO ")
      SqlOperadora.Add("   AND DATAEXPORTACAO IS NOT NULL  ")
      SqlOperadora.ParamByName("OPERADORA").Value = Voperadora
      SqlOperadora.ParamByName("TABTIPO").Value = Sqltipo.FieldByName("tabtipo").AsInteger
      SqlOperadora.Active = True

      If SqlOperadora.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
        bsShowMessage("Data de adesão do módulo é menor ou igual a data final do envio de beneficiários a ANS.", "E")
        CanContinue = False
        Exit Sub
      End If
    Else
      Voperadora = Sqltipo.FieldByName("ADMINISTRADORA").AsInteger

      SqlOperadora.Active = False
      SqlOperadora.Clear
      SqlOperadora.Add("SELECT MAX(DATAFINAL)DATAFINAL      ")
      SqlOperadora.Add("  FROM SAM_MSPROCESSO               ")
      SqlOperadora.Add(" WHERE OPERADORAADM = :OPERADORA    ")
      SqlOperadora.Add("   AND SITUACAO  = 'P'              ")
      SqlOperadora.Add("   AND TABTIPOOPERADORA = :TABTIPO  ")
      SqlOperadora.Add("   AND DATAEXPORTACAO IS NOT NULL   ")
      SqlOperadora.ParamByName("OPERADORA").Value = Voperadora
      SqlOperadora.ParamByName("TABTIPO").Value = Sqltipo.FieldByName("tabtipo").AsInteger
      SqlOperadora.Active = True

      If SqlOperadora.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
        bsShowMessage("Data de adesão do módulo é menor ou igual a data final do envio de beneficiários a ANS.", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
  Else
    Sqltipo.Active = False
    Sqltipo.Clear
    Sqltipo.Add("SELECT MAX(B.COMPETENCIA) COMPETENCIA")
    Sqltipo.Add("  FROM GER_BENEF_COMPET B,           ")
    Sqltipo.Add("       GER_BENEF_COMPETRESUMO BC,    ")
    Sqltipo.Add("       SAM_CONTRATO C                ")
    Sqltipo.Add(" WHERE B.HANDLE =  BC.COMPETENCIA    ")
    Sqltipo.Add("   AND BC.CONTRATO = C.HANDLE        ")
    Sqltipo.Add("   AND C.HANDLE = :CONTRATO          ")
    Sqltipo.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
    Sqltipo.Active = True

    If Sqltipo.FieldByName("COMPETENCIA").AsDateTime <>0 Then

      vCompFinal = UltimoDiaCompetencia(Sqltipo.FieldByName("COMPETENCIA").AsDateTime)

      If vCompFinal >= CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
        bsShowMessage("Data de adesão do módulo é menor ou igual a competencia final da totalização dos beneficiarios.", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

  End If

End If

Set Sqltipo = Nothing
Set SqlOperadora = Nothing
'FIM ANDERSON
'--------------------------------------------------------------------------------------------------------------------------

If CurrentQuery.FieldByName("VERIFICADEPENDENTEAGREGADO").AsString = "S" Then
  If(CurrentQuery.FieldByName("ACEITATITULAR").AsString = "N")Then
  bsShowMessage("Para verificar dependentes agregados, o campo Aceita Titular deve estar ativo!", "E")
  CanContinue = False
End If
If(CurrentQuery.FieldByName("ACEITADEPENDENTES").AsString = "S")Then
bsShowMessage("Para verificar dependentes agregados, o campo Aceita Dependente deve estar inativo!", "E")
CanContinue = False
End If
If(CurrentQuery.FieldByName("ACEITAAGREGADOS").AsString = "S")Then
bsShowMessage("Para verificar dependentes agregados, o campo Aceita Agregados deve estar inativo!", "E")
CanContinue = False
End If
If(CurrentQuery.FieldByName("AUTOMATICO").AsString = "S")Then
bsShowMessage("Para verificar dependentes agregados, o campo Automático deve estar inativo!", "E")
CanContinue = False
End If
If(CurrentQuery.FieldByName("PROPAGAR").AsString = "S")Then
bsShowMessage("Para verificar dependentes agregados, o campo Propagar deve estar inativo!", "E")
CanContinue = False
End If
End If

'Dim qAceitaDependente As Object
'Set qAceitaDependente =NewQuery

'Juliano -07/08/2003
'Não é necessário fazer esta checagem porque o contrato individual pode ter dependente,desde que não tenha titular
'If CurrentQuery.FieldByName("ACEITADEPENDENTES").AsString ="S" Then
' qAceitaDependente.Add("SELECT * FROM SAM_CONTRATO WHERE HANDLE= :HCONTRATO")
' qAceitaDependente.ParamByName("HCONTRATO").AsInteger =CurrentQuery.FieldByName("CONTRATO").AsInteger
' qAceitaDependente.Active=True

' If qAceitaDependente.FieldByName("TABTIPOCONTRATO").AsString ="3" Then 'Individual
'    MsgBox("Plano Individual, não aceita módulos dependentes!")
'    CanContinue =False
' End If
'End If
'Set qAceitaDependente=Nothing

If CurrentQuery.FieldByName("SEGUNDAPARCELA").AsString = "2" And _
                            CurrentQuery.FieldByName("PRIMEIRAPARCELA").AsString <>"2" Then
  CanContinue = False
  bsShowMessage("Para segunda parcela 'Proporcional' a primeira parcela deve ser integral", "E")
  Exit Sub
End If

Dim SQL As Object

Set SQL = NewQuery

SQL.Clear
SQL.Add("SELECT NAOREGISTRARNOMS,TABTIPOCONTRATO,TIPOCONTRATACAOEMPRESARIAL")
SQL.Add("FROM SAM_CONTRATO")
SQL.Add("WHERE HANDLE = :HCONTRATO")
SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
SQL.Active = True

Dim vTabTipoContrato As Integer
Dim vTipoContratacaoEmpresarial As String

vTabTipoContrato = SQL.FieldByName("TABTIPOCONTRATO").AsInteger
vTipoContratacaoEmpresarial = SQL.FieldByName("TIPOCONTRATACAOEMPRESARIAL").AsString

If(SQL.FieldByName("NAOREGISTRARNOMS").AsString = "N")And _
   (CurrentQuery.FieldByName("OBRIGATORIO").AsString = "S")And _
   (CurrentQuery.FieldByName("REGISTROMS").IsNull)Then
bsShowMessage("O registro no Ministério da Saúde é obrigatório para módulos obrigatórios", "E")
CanContinue = False
End If

If Not CurrentQuery.FieldByName("REGISTROMS").IsNull Then
  SQL.Clear
  SQL.Add("SELECT DATAVENCIMENTO,TIPOCONTRATACAO,NOVAREGULAMENTACAO")
  SQL.Add("FROM SAM_REGISTROMS")
  SQL.Add("WHERE HANDLE = :HREGISTROMS")
  SQL.ParamByName("HREGISTROMS").Value = CurrentQuery.FieldByName("REGISTROMS").AsInteger
  SQL.Active = True

  If Not SQL.FieldByName("DATAVENCIMENTO").IsNull And _
                         (SQL.FieldByName("DATAVENCIMENTO").AsDateTime <ServerDate)Then
    bsShowMessage("O Registro no Ministério da Saúde está vencido! Verifique", "E")
  End If

  If CurrentQuery.State = 3 Then
    If SQL.FieldByName("NOVAREGULAMENTACAO").AsString = "S" Then
      If vTabTipoContrato = 1 Then 'Empresarial
        If vTipoContratacaoEmpresarial = "E" And SQL.FieldByName("TIPOCONTRATACAO").AsString <>"1" Then
          bsShowMessage("Registro do Ministério não permitido para contrato com tipo de contratação 'Coletivo Empresarial'", "E")
          CanContinue = False
          Exit Sub
        End If

        If vTipoContratacaoEmpresarial = "A" And SQL.FieldByName("TIPOCONTRATACAO").AsString <>"2" Then
          bsShowMessage("Registro do Ministério não permitido para contrato com tipo de contratação 'Coletivo por Adesão'", "E")
          CanContinue = False
          Exit Sub
        End If
      Else
        If SQL.FieldByName("TIPOCONTRATACAO").AsString <>"3" Then
          If vTabTipoContrato = 2 Then
            bsShowMessage("Tipo de Contratação do Registro no Ministério da Saúde não permitido para contrato com tipo de 'Familiar'", "E")
            CanContinue = False
            Exit Sub
          Else
            bsShowMessage("Tipo de Contratação do Registro no Ministério da Saúde não permitido para contrato 'Individual'", "E")
            CanContinue = False
            Exit Sub
          End If
        End If
      End If
    End If
  End If
End If


'  Set SQL =Nothing

vEstadoTabela = CurrentQuery.State
'  Dim SQL As Object
'  Set SQL=NewQuery
SQL.Clear
SQL.Add("SELECT DATAADESAO, TIPOFATURAMENTO FROM SAM_CONTRATO")
SQL.Add("WHERE HANDLE = :CONTRATO")
SQL.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
SQL.Active = True
If CurrentQuery.FieldByName("DATAADESAO").AsDateTime <SQL.FieldByName("DATAADESAO").AsDateTime Then
  BsShowMessage("Data de adesão inválida!", "E")
  CanContinue = False
Else
  If(CurrentQuery.FieldByName("PRIMEIRAPARCELA").AsString <>"3")And _
     (CurrentQuery.FieldByName("SEGUNDAPARCELA").AsString = "1")Then ' isento ou integral
  If Not CurrentQuery.FieldByName("PARCELADIAS").IsNull Then
    bsShowMessage("Informar quantidade de dias somente para primeira/segunda parcela proporcional!", "E")
    CanContinue = False
    Exit Sub
  End If
End If
End If

SQL.Active = False
Set SQL = Nothing
If CanContinue = True Then
  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String


  Condicao = "AND MODULO = " + CurrentQuery.FieldByName("MODULO").AsString

  'sms 49923
  If (CurrentQuery.FieldByName("REGISTROMS").IsNull) Then
    Condicao = Condicao + " AND REGISTROMS IS NULL"
  Else
    Condicao = Condicao + " AND REGISTROMS = " + CurrentQuery.FieldByName("REGISTROMS").AsString
  End If

  Condicao = Condicao + " AND PLANO = " + CurrentQuery.FieldByName("PLANO").AsString


  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_CONTRATO_MOD", "DATAADESAO", "DATACANCELAMENTO", CurrentQuery.FieldByName("DATAADESAO").AsDateTime, CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime, "CONTRATO", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
  End If
End If

If CurrentQuery.FieldByName("PRECOPORVALOROUCOTA").Value = "C" And _
                            Not(VerificaSeRateio)Then
  CanContinue = False
  bsShowMessage("Preço do módulo por cota só é permitido para Contratos de Custo Operacional - Rateio", "E")
End If

'sms 22962
Dim qSQL As Object
Set qSQL = NewQuery

If CurrentQuery.FieldByName("TABFRANQUIAURGENCIAEMERGENCIA").AsInteger = 2 Then 'Checa se o Contrato é Pré-pagamento ou Autogestão código 120 e 130 POIS ESTÁ MARCADO QUE POSSUI Frq/Urg/Emerg


  Dim vTipoFaturamento As Integer

  Dim qSQL3 As Object
  Set qSQL3 = NewQuery
  qSQL3.Clear
  qSQL3.Add("SELECT A.CODIGO ")
  qSQL3.Add("FROM SAM_CONTRATO C, SIS_TIPOFATURAMENTO A ")
  qSQL3.Add("WHERE C.HANDLE = :HCONTRATO AND A.HANDLE = C.TIPOFATURAMENTO ")
  qSQL3.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qSQL3.Active = True
  vTipoFaturamento = qSQL3.FieldByName("CODIGO").AsInteger
  Set qSQL3 = Nothing


  If vTipoFaturamento = 130 Then 'Auto-gestão
    qSQL.Clear
    qSQL.Add("    SELECT HANDLE ")
    qSQL.Add("    FROM SAM_CONTRATO_AUTOGESTAO ")
    qSQL.Add("    WHERE PERMITESUPLEMENTACAOPF = 'S' ")
    qSQL.Add("      AND CONTRATO = :HCONTRATO ")
    qSQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
    qSQL.Active = True
    If Not qSQL.EOF Then 'Permite suplementação  => só Frq/Urg/Emerg em módulo suplementar
      Dim qSQL2 As Object
      Set qSQL2 = NewQuery
      qSQL2.Clear
      qSQL2.Add("    SELECT TIPOMODULO ")
      qSQL2.Add("    FROM SAM_MODULO ")
      qSQL2.Add("    WHERE HANDLE = :HMODULO ")
      qSQL2.ParamByName("HMODULO").Value = CurrentQuery.FieldByName("MODULO").AsInteger
      qSQL2.Active = True
      If qSQL2.FieldByName("TIPOMODULO").AsString <> "S" Then
        bsShowMessage("Em contratos com suplementação de PF, apenas módulos de suplementação podem possuir franquia de urgência/emergência", "E")
        CanContinue = False
        Set qSQL2 = Nothing
        Exit Sub
      End If
      Set qSQL2 = Nothing
    Else 'Não permite suplementação => só Frq/Urg/Emerg em módulo Cobertura
      Dim qSQL4 As Object
      Set qSQL4 = NewQuery
      qSQL4.Clear
      qSQL4.Add("    SELECT TIPOMODULO ")
      qSQL4.Add("    FROM SAM_MODULO ")
      qSQL4.Add("    WHERE HANDLE = :HMODULO ")
      qSQL4.ParamByName("HMODULO").Value = CurrentQuery.FieldByName("MODULO").AsInteger
      qSQL4.Active = True
      If qSQL4.FieldByName("TIPOMODULO").AsString <> "C" Then
        bsShowMessage("Em contratos sem suplementação de PF, apenas módulos de cobertura podem possuir franquia de urgência/emergência", "E")
        CanContinue = False
        Set qSQL4 = Nothing
        Exit Sub
      End If
      Set qSQL4 = Nothing
    End If
  ElseIf vTipoFaturamento = 120 Then 'PRÉ-PAGAMENTO => só modulo de cobertura
    Dim qSQL5 As Object
    Set qSQL5 = NewQuery
    qSQL5.Clear
    qSQL5.Add("    SELECT TIPOMODULO ")
    qSQL5.Add("    FROM SAM_MODULO ")
    qSQL5.Add("    WHERE HANDLE = :HMODULO ")
    qSQL5.ParamByName("HMODULO").Value = CurrentQuery.FieldByName("MODULO").AsInteger
    qSQL5.Active = True
    If qSQL5.FieldByName("TIPOMODULO").AsString <> "C" Then
      bsShowMessage("Em contratos de pré-pagamento, apenas módulos de cobertura podem possuir franquia de urgência/emergência", "E")
      CanContinue = False
      Set qSQL5 = Nothing
      Exit Sub
    End If
    Set qSQL5 = Nothing
  Else
    bsShowMessage("Tipo de faturamento do contrato não permite franquia de urgência/emergência", "E")
    CanContinue = False
    Exit Sub
  End If
End If

Set qSQL = Nothing

'fim sms 22962

'Verifica se a data de cancelamento ou data de adesão é inferior a data de fechamento
'Rodrigo -14/01/2003
Dim qDataFechamento As Object
Set qDataFechamento = NewQuery
qDataFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
qDataFechamento.Active = True

If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
  If qDataFechamento.FieldByName("DATAFECHAMENTO").AsDateTime >CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime Then
    bsShowMessage("Não é possível cancelar módulo com data de cancelamento inferior a data de de fechamento - Parâmetros Gerais", "E")
    CanContinue = False
  End If
End If

If Not CurrentQuery.FieldByName("DATAADESAO").IsNull And CurrentQuery.State = 3 Then
  If qDataFechamento.FieldByName("DATAFECHAMENTO").AsDateTime >CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Não é possível cancelar módulo com data de cancelamento inferior a data de de fechamento - Parâmetros Gerais", "E")
    CanContinue = False
  End If
End If

Set qDataFechamento = Nothing

Dim qVerificarAdesaoPlano As Object
Set qVerificarAdesaoPlano = NewQuery
qVerificarAdesaoPlano.Add("SELECT DATAADESAO          ")
qVerificarAdesaoPlano.Add("  FROM SAM_CONTRATO_PLANO    ")
qVerificarAdesaoPlano.Add(" WHERE PLANO    = :pPLANO    ")
qVerificarAdesaoPlano.Add("   AND CONTRATO = :pCONTRATO ")
qVerificarAdesaoPlano.ParamByName("pPLANO"   ).AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
qVerificarAdesaoPlano.ParamByName("pCONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
qVerificarAdesaoPlano.Active = True

If qVerificarAdesaoPlano.FieldByName("DATAADESAO").AsDateTime > CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
  bsShowMessage("Data de adesão do módulo é menor que data de adesão do plano no contrato.", "E")
  CanContinue = False
  Exit Sub
End If

Set qVerificarAdesaoPlano = Nothing

Dim qVerificaModuloCadastradoObrigatorio As Object
Set qVerificaModuloCadastradoObrigatorio = NewQuery
Dim qModuloEraObrigatorio As Object
Set qModuloEraObrigatorio = NewQuery

qVerificaModuloCadastradoObrigatorio.Clear
qVerificaModuloCadastradoObrigatorio.Add("SELECT COUNT(1) QTDE                                  ")
qVerificaModuloCadastradoObrigatorio.Add("  FROM SAM_BENEFICIARIO B                             ")
qVerificaModuloCadastradoObrigatorio.Add(" WHERE EXISTS (SELECT 1                           ")
qVerificaModuloCadastradoObrigatorio.Add("                 FROM SAM_BENEFICIARIO_MOD        ")
qVerificaModuloCadastradoObrigatorio.Add("                WHERE BENEFICIARIO = B.HANDLE     ")
qVerificaModuloCadastradoObrigatorio.Add("                  AND MODULO = :HANDLE)           ")
qVerificaModuloCadastradoObrigatorio.Add("   AND CONTRATO = :CONTRATO                           ")
qVerificaModuloCadastradoObrigatorio.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
qVerificaModuloCadastradoObrigatorio.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
qVerificaModuloCadastradoObrigatorio.Active = True

qModuloEraObrigatorio.Clear
qModuloEraObrigatorio.Add("SELECT HANDLE, OBRIGATORIO FROM SAM_cONTRATO_MOD WHERE HANDLE = :HANDLE")
qModuloEraObrigatorio.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
qModuloEraObrigatorio.Active = True

If CurrentQuery.FieldByName("OBRIGATORIO").AsString = "S" Then
  If (qModuloEraObrigatorio.FieldByName("OBRIGATORIO").AsString = "N") Then
    If qVerificaModuloCadastradoObrigatorio.FieldByName("QTDE").AsInteger > 0 Then
      bsShowMessage("Módulo não pode ser alterado para obrigatório, pois já existem beneficiários cadastrados com este módulo.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If
  If CurrentQuery.FieldByName("PROPAGAR").AsString = "S" Then
    bsShowMessage("Em módulos obrigatórios, não é permitido marcar o flag Propagar.","E")
    CanContinue = False
    Exit Sub
  End If
Else
  If (qModuloEraObrigatorio.FieldByName("OBRIGATORIO").AsString = "S") Then
    If qVerificaModuloCadastradoObrigatorio.FieldByName("QTDE").AsInteger > 0 Then
      bsShowMessage("Módulo não pode ser alterado para opcional, pois já existem beneficiários cadastrados com este módulo.", "E")
      CanContinue = False
      Exit Sub
  End If
  End If
End If

Set qVerificaModuloCadastradoObrigatorio = Nothing
Set qModuloEraObrigatorio = Nothing

End Sub

Public Function VerificaSeRateio As Boolean
  Dim SQLRotFin As Object
  Set SQLRotFin = NewQuery
  SQLRotFin.Add("SELECT B.CODIGO FROM SAM_CONTRATO A, SIS_TIPOFATURAMENTO B")
  SQLRotFin.Add("WHERE A.HANDLE = :HANDLE")
  SQLRotFin.Add("  AND B.HANDLE = A.TIPOFATURAMENTO")
  SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("CONTRATO").Value
  SQLRotFin.Active = True
  If SQLRotFin.FieldByName("CODIGO").Value = 140 Then
    VerificaSeRateio = True
  Else
    VerificaSeRateio = False
  End If
  SQLRotFin.Active = False
  Set SQLRotFin = Nothing
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOBENEFICIARIOSATIVOS"
			BOTAOBENEFICIARIOSATIVOS_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPROPAGAR"
			BOTAOPROPAGAR_OnClick
		Case "BOTAOREATIVAR"
			BOTAOREATIVAR_OnClick
		Case "BOTAOTRANSFEREMODULO"
			BOTAOTRANSFEREMODULO_OnClick
	End Select
End Sub
