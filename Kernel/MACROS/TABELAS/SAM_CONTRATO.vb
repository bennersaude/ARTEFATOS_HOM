'HASH: C2A0719D9CAE8B9849E3B508A79472AD

'Macro: SAM_CONTRATO
' Leonardo -08/12/2000 -Inserir registros referentes a SAM_PLANOS_EVENTOSEMAUTOR para os contratos novos
'Última alteração: Milton/17/01/2002 -SMS 5976
'#Uses "*bsShowMessage"

Option Explicit
Dim EstadoTabela As Long
Dim VDataUltimoReajuste As Date
Dim vCobrancaDeEvento As String
Dim vNaoRegistrarNoMS As String
Dim vDiaCobranca As Integer
Dim vNumeroFamiliaAutomaticoAnterior As String
Dim vNumeroBenefAutomaticoAnterior As String
Dim vMotivoBloqueioAutomatico As Integer
Dim vsModoEdicao            As String
Dim vsXMLContainerEnderecos As String
Dim vsXMLEnderecosExcluidos As String

Public Sub HerdarRelacTpDepEEstadoCivil(pHandleContrato As Long)
	Dim qAux1 As Object
	Dim qAux2 As Object
	Dim vNewHandle As Long
	Set qAux1 = NewQuery
	Set qAux2 = NewQuery

	'Herdando relacionamento entre dependentes - inicio
		'Preparando insert na tabela SAM_CONTRATO_TPDEP_NAOPERMITID
		qAux2.Add("INSERT INTO SAM_CONTRATO_TPDEP_NAOPERMITID            ")
		qAux2.Add("  ( HANDLE, CONTRATOTPDEP, TPDEPNAOPERMITIDO)         ")
		qAux2.Add("  VALUES                                              ")
		qAux2.Add("  (:HANDLE,:CONTRATOTPDEP,:TIPODEPENDENTENAOPERMITIDO)")

		'Buscando configuração a partir do plano do contrato
		qAux1.Add("SELECT CTD.HANDLE CONTRATOTPDEP, DNP.TPDEPNAOPERMITIDO TIPODEPENDENTENAOPERMITIDO ")
		qAux1.Add("  FROM SAM_CONTRATO                 CON,                                ")
		qAux1.Add("       SAM_PLANO                    PLA,                                ")
		qAux1.Add("       SAM_PLANO_TPDEP              PTD,                                ")
		qAux1.Add("       SAM_PLANO_TPDEP_NAOPERMITIDO DNP,                                ")
		qAux1.Add("       SAM_TIPODEPENDENTE           TDP,                                ")
		qAux1.Add("       SAM_CONTRATO_TPDEP           CTD                                 ")
		qAux1.Add(" WHERE CON.HANDLE = :CONTRATO                                           ")
		qAux1.Add("   AND (PLA.HANDLE = CON.PLANO          )                               ")
		qAux1.Add("   AND (PLA.HANDLE = PTD.PLANO          )                               ")
		qAux1.Add("   AND (PTD.HANDLE = DNP.PLANOTPDEP     )                               ")
		qAux1.Add("   AND (TDP.HANDLE = PTD.TIPODEPENDENTE )                               ")
		qAux1.Add("   AND (CON.HANDLE = CTD.CONTRATO AND TDP.HANDLE = CTD.TIPODEPENDENTE ) ")
		qAux1.ParamByName("CONTRATO").AsInteger = pHandleContrato
		qAux1.Active = True

		'Herdando efetivamente o relacionamento entre dependentes
		qAux1.First
		While Not qAux1.EOF
		    vNewHandle = NewHandle("SAM_CONTRATO_TPDEP_NAOPERMITID")
			qAux2.Active = False
			qAux2.ParamByName("HANDLE").AsInteger                     = vNewHandle
			qAux2.ParamByName("CONTRATOTPDEP").AsInteger              = qAux1.FieldByName("CONTRATOTPDEP").AsInteger
			qAux2.ParamByName("TIPODEPENDENTENAOPERMITIDO").AsInteger = qAux1.FieldByName("TIPODEPENDENTENAOPERMITIDO").AsInteger
			On Error GoTo EXCEPT1
  				qAux2.ExecSQL
  			EXCEPT1:
				qAux1.Next
		Wend
	'Herdando relacionamento entre dependentes - fim

	're-inicializa queries
	qAux1.Active = False
	qAux1.Clear
	qAux2.Active = False
	qAux2.Clear


	'Herdando configuração de estado civil do titular - inicio
		'Preparando insert na tabela SAM_CONTRATO_TPDEP_ESTCIVILTIT
		qAux2.Add("INSERT INTO SAM_CONTRATO_TPDEP_ESTCIVILTIT ")
		qAux2.Add("  ( HANDLE, CONTRATOTPDEP, ESTADOCIVIL)    ")
		qAux2.Add("  VALUES                                   ")
		qAux2.Add("  (:HANDLE,:CONTRATOTPDEP,:ESTADOCIVIL)    ")

		'Buscando configuração a partir do plano do contrato
		qAux1.Add("SELECT CTD.HANDLE CONTRATOTPDEP, DNP.ESTADOCIVIL                       ")
		qAux1.Add("  FROM SAM_CONTRATO                   CON,                             ")
		qAux1.Add("       SAM_PLANO                      PLA,                             ")
		qAux1.Add("       SAM_PLANO_TPDEP                PTD,                             ")
		qAux1.Add("       SAM_PLANO_TPDEP_ESTADOCIVILTIT DNP,                             ")
		qAux1.Add("       SAM_TIPODEPENDENTE             TDP,                             ")
		qAux1.Add("       SAM_CONTRATO_TPDEP             CTD                              ")
		qAux1.Add(" WHERE CON.HANDLE = :CONTRATO                                          ")
		qAux1.Add("   AND (PLA.HANDLE = CON.PLANO          )                              ")
		qAux1.Add("   AND (PLA.HANDLE = PTD.PLANO          )                              ")
		qAux1.Add("   AND (PTD.HANDLE = DNP.PLANOTPDEP     )                              ")
		qAux1.Add("   AND (TDP.HANDLE = PTD.TIPODEPENDENTE )                              ")
		qAux1.Add("   AND (CON.HANDLE = CTD.CONTRATO AND TDP.HANDLE = CTD.TIPODEPENDENTE )")
		qAux1.ParamByName("CONTRATO").AsInteger = pHandleContrato
		qAux1.Active = True

		'Herdando efetivamente os estados civis permitidos do que permitem a inserção do tipo de beneficiario
		qAux1.First
		While Not qAux1.EOF
		    vNewHandle = NewHandle("SAM_CONTRATO_TPDEP_ESTCIVILTIT")
			qAux2.Active = False
			qAux2.ParamByName("HANDLE").AsInteger        = vNewHandle
			qAux2.ParamByName("CONTRATOTPDEP").AsInteger = qAux1.FieldByName("CONTRATOTPDEP").AsInteger
			qAux2.ParamByName("ESTADOCIVIL").AsInteger   = qAux1.FieldByName("ESTADOCIVIL").AsInteger
			On Error GoTo EXCEPT2
  				qAux2.ExecSQL
  			EXCEPT2:
				qAux1.Next
		Wend
	'Herdando configuração de estado civil do titular - fim

	're-inicializa queries
	qAux1.Active = False
	qAux1.Clear
	qAux2.Active = False
	qAux2.Clear


	'Herdando configuração de estado civil dos dependentes - inicio
		'Preparando insert na tabela SAM_CONTRATO_TPDEP_ESTCIVILDEP
		qAux2.Add("INSERT INTO SAM_CONTRATO_TPDEP_ESTCIVILDEP ")
		qAux2.Add("  ( HANDLE, CONTRATOTPDEP, ESTADOCIVIL)    ")
		qAux2.Add("  VALUES                                   ")
		qAux2.Add("  (:HANDLE,:CONTRATOTPDEP,:ESTADOCIVIL)    ")

		'Buscando configuração a partir do plano do contrato
		qAux1.Add("SELECT CTD.HANDLE CONTRATOTPDEP, DNP.ESTADOCIVIL                       ")
		qAux1.Add("  FROM SAM_CONTRATO                   CON,                             ")
		qAux1.Add("       SAM_PLANO                      PLA,                             ")
		qAux1.Add("       SAM_PLANO_TPDEP                PTD,                             ")
		qAux1.Add("       SAM_PLANO_TPDEP_ESTADOCIVILDEP DNP,                             ")
		qAux1.Add("       SAM_TIPODEPENDENTE             TDP,                             ")
		qAux1.Add("       SAM_CONTRATO_TPDEP             CTD                              ")
		qAux1.Add(" WHERE CON.HANDLE = :CONTRATO                                          ")
		qAux1.Add("   AND (PLA.HANDLE = CON.PLANO          )                              ")
		qAux1.Add("   AND (PLA.HANDLE = PTD.PLANO          )                              ")
		qAux1.Add("   AND (PTD.HANDLE = DNP.PLANOTPDEP     )                              ")
		qAux1.Add("   AND (TDP.HANDLE = PTD.TIPODEPENDENTE )                              ")
		qAux1.Add("   AND (CON.HANDLE = CTD.CONTRATO AND TDP.HANDLE = CTD.TIPODEPENDENTE )")
		qAux1.ParamByName("CONTRATO").AsInteger = pHandleContrato
		qAux1.Active = True

		'Herdando efetivamente os estados civis permitidos do que permitem a inserção do tipo de beneficiario
		qAux1.First
		While Not qAux1.EOF
		    vNewHandle = NewHandle("SAM_CONTRATO_TPDEP_ESTCIVILDEP")
			qAux2.Active = False
			qAux2.ParamByName("HANDLE").AsInteger        = vNewHandle
			qAux2.ParamByName("CONTRATOTPDEP").AsInteger = qAux1.FieldByName("CONTRATOTPDEP").AsInteger
			qAux2.ParamByName("ESTADOCIVIL").AsInteger   = qAux1.FieldByName("ESTADOCIVIL").AsInteger
			On Error GoTo EXCEPT3
  				qAux2.ExecSQL
  			EXCEPT3:
				qAux1.Next
		Wend
	'Herdando configuração de estado civil do titular - fim



End Sub


Public Function GeraRegContratoCartaoMotivo()
  '
  GeraRegContratoCartaoMotivo = True

  Dim vlHAlteracao As Long
  Dim vlHSegVia As Long
  Dim vlValorSegVia As Double
  Dim vlValorAlteracaoCad As Double

  Dim QSP As Object
  Set QSP = NewQuery

  QSP.Add("SELECT CARTAOMOTIVOSEGVIA,VALORSEGVIA, MOTIVOALTERACAOCADASTRAL, VALORALTERACAOCADASTRAL FROM SAM_PARAMETROSBENEFICIARIO")
  QSP.Active = True

  If QSP.EOF Then
    bsShowMessage("Tabela de paramentros Invalida - Processo Abortado", "E")
    Exit Function
  End If

  vlHAlteracao = QSP.FieldByName("MOTIVOALTERACAOCADASTRAL").AsInteger
  vlValorAlteracaoCad = QSP.FieldByName("VALORALTERACAOCADASTRAL").AsFloat
  vlHSegVia = QSP.FieldByName("CARTAOMOTIVOSEGVIA").AsInteger
  vlValorSegVia = QSP.FieldByName("VALORSEGVIA").AsFloat

  Set QSP = Nothing

  Dim QSU As Object
  Set QSU = NewQuery

  QSU.Add("SELECT HANDLE FROM SAM_CONTRATO_CARTAOMOTIVO WHERE CONTRATO=:CONTRATO AND CARTAOMOTIVO = :CARTAOMOTIVO")
  QSU.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QSU.ParamByName("CARTAOMOTIVO").AsInteger = vlHAlteracao
  QSU.Active = True

  If QSU.EOF Then

    Dim Qu As Object

    Set Qu = NewQuery

    Qu.Clear
    Qu.Add("INSERT INTO SAM_CONTRATO_CARTAOMOTIVO (HANDLE,CONTRATO,CARTAOMOTIVO,COBRAREMISSAO,TAXACARTAO) VALUES (:HANDLE,:CONTRATO,:CARTAOMOTIVO,:COBRAREMISSAO,:TAXACARTAO)")
    Qu.ParamByName("HANDLE").Value = NewHandle("SAM_CONTRATO_CARTAOMOTIVO")
    Qu.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("HANDLE").Value
    Qu.ParamByName("CARTAOMOTIVO").Value = vlHAlteracao

    If vlValorAlteracaoCad >0 Then
      Qu.ParamByName("COBRAREMISSAO").Value = "S"
      Qu.ParamByName("TAXACARTAO").Value = vlValorAlteracaoCad
    Else
      Qu.ParamByName("COBRAREMISSAO").Value = "N"
      Qu.ParamByName("TAXACARTAO").Value = 0
    End If


    Qu.ExecSQL


    Set Qu = Nothing

  End If

  Set QSU = Nothing

  Dim QSS As Object
  Set QSS = NewQuery

  QSS.Add("SELECT HANDLE FROM SAM_CONTRATO_CARTAOMOTIVO WHERE CONTRATO=:CONTRATO AND CARTAOMOTIVO = :CARTAOMOTIVO")
  QSS.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  QSS.ParamByName("CARTAOMOTIVO").AsInteger = vlHSegVia

  QSS.Active = True

  If QSS.EOF Then

    Dim Qu2 As Object
    Set Qu2 = NewQuery

    Qu2.Clear
    Qu2.Add("INSERT INTO SAM_CONTRATO_CARTAOMOTIVO (HANDLE,CONTRATO,CARTAOMOTIVO,COBRAREMISSAO,TAXACARTAO) VALUES (:HANDLE,:CONTRATO,:CARTAOMOTIVO,:COBRAREMISSAO,:TAXACARTAO)")
    Qu2.ParamByName("HANDLE").Value = NewHandle("SAM_CONTRATO_CARTAOMOTIVO")
    Qu2.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("HANDLE").Value
    Qu2.ParamByName("CARTAOMOTIVO").Value = vlHSegVia

    If vlValorSegVia >0 Then
      Qu2.ParamByName("COBRAREMISSAO").Value = "S"
      Qu2.ParamByName("TAXACARTAO").Value = vlValorSegVia
    Else
      Qu2.ParamByName("COBRAREMISSAO").Value = "N"
      Qu2.ParamByName("TAXACARTAO").Value = 0
    End If


    Qu2.ExecSQL


    Set Qu2 = Nothing

  End If

  Set QSS = Nothing

  GeraRegContratoCartaoMotivo = False

End Function

Public Sub BOTAOADICIONARPLANO_OnClick()
  Dim DLLADICIONARPLANO As Object
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Dim vsTabelasPreco As String
  Dim vsMensagemRetorno As String
  Dim viRetorno As Integer

  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If (BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    0, _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)) Then
    bsShowMessage("Não é permitido adicionar plano ao contrato por motivo de suspensão!", "I")
    Exit Sub
  End If
  Set BSBen001Dll = Nothing

  If (VisibleMode) Then
    If (CurrentQuery.State <> 1) Then
	  bsShowMessage("O registro não pode estar em edição", "I")
	  Exit Sub
	End If

    If Not (CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull) Then
	  bsShowMessage("O contrato já está cancelado !", "I")
	  Exit Sub
	End If

    Set DLLADICIONARPLANO = CreateBennerObject("BSINTERFACE0068.ROTINAS")

    viRetorno = DLLADICIONARPLANO.AdicionarPlano(CurrentQuery.FieldByName("HANDLE").AsInteger)
  End If

  Set DLLADICIONARPLANO = Nothing
End Sub

Public Sub BOTAOALTERAPADRAOPRECO_OnClick()
  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Dim Interface As Object

  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    0, _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido alterar o padrão de preço por motivo de suspensão!", "E")
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------

  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição","I")
    Exit Sub
  End If

  If bsShowMessage("Confirma a alteração do padrão de preço do contrato para a família?", "Q") = vbYes Then

    Set Interface = CreateBennerObject("CONTRATO.ContratoInterface")
    Interface.AlteraPadraoPreco(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Set Interface = Nothing

    WriteAudit("A", HandleOfTable("SAM_CONTRATO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Padrão de preço do módulo alterado para ser na Família")

    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOALTERARADESAO_OnClick()
  Dim vcContainer     As Object
  Dim BSINTERFACE0002 As Object
  Dim vsMensagem      As String
  Dim viRetorno       As Long

  Set vcContainer = NewContainer
  Set BSINTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

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

Public Sub BOTAOBENEFATIVOS_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT COUNT(*) NRECS FROM SAM_BENEFICIARIO B     ")
  SQL.Add("WHERE B.CONTRATO=:HCONTRATO                       ")
  SQL.Add("  AND (   B.DATACANCELAMENTO IS NULL              ")
  SQL.Add("       OR B.DATACANCELAMENTO >= :HOJE             ")
  SQL.Add("       OR B.ATENDIMENTOATE   >= :HOJE)            ")
  SQL.Add("  AND B.DATAADESAO <= :HOJE                       ") 'SMS 81279 - Marcelo Barbosa - 08/05/2007
  SQL.Add("  AND B.DATABLOQUEIO IS NULL                      ")


  SQL.ParamByName("HCONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("HOJE").Value = ServerDate

  SQL.Active = True
  bsShowMessage("Beneficiários ativos:" + Str(SQL.FieldByName("NRECS").AsInteger), "I")

  SQL.Active = False
  Set SQL = Nothing
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  Dim Interface As Object
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
	Dim vsMensagemErro As String
	Dim viRetorno As Integer
    Dim vvContainer As CSDContainer

	Set vvContainer = NewContainer

	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


	viRetorno = Interface.Exec(CurrentSystem, _
    						   1, _
                               "TV_FORM0014", _
                               "Cancelamento de Contrato", _
       	                       0, _
           	                   180, _
               	               420, _
                   	           False, _
                       	       vsMensagemErro, _
                           	   vvContainer)

	Select Case viRetorno
      Case -1
	  	bsShowMessage("Operação cancelada pelo usuário!", "I")
  	  Case  1
   	  	bsShowMessage(vsMensagemErro, "I")
	End Select

    Set Interface = Nothing
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  Else
	bsShowMessage("Contrato já Cancelado!","I")
  End If
End Sub


Public Sub BOTAODIACOBRANCA_OnClick()
  If CurrentQuery.State <>1 Then
    bsShowMessage("A tabela não pode estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger <>1 Or _
                              CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "C" Then
    bsShowMessage("O contrato deve ser EMPRESARIAL e o local de faturamento deve ser na FAMÍLIA", "I")
    Exit Sub
  End If

  Dim Interface As Object

  Set Interface = CreateBennerObject("SamFaturamento.Faturamento")
  Interface.AlterarDiaCobranca(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing

End Sub

Public Sub BOTAOENDERECO_OnClick()
	Dim Interface As Object
	If Not WebMode Then
		If TABTIPOCONTRATO.PageIndex <> 0 Then
			bsShowMessage("O Tipo de Contrato não possibilita inclusão de Endereço", "I")
			Exit Sub
		End If

		Dim viHEnderecoCorrespondencia As Long
		viHEnderecoCorrespondencia = CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger

		Dim dllBSInterface0028 As Object
		Set dllBSInterface0028 = CreateBennerObject("BSInterface0028.Endereco")

		Dim msgErro As String
		Dim numErro As Long

		On Error GoTo Except
			Dim vHandleContrato As Long
			Dim vsMensagem As String

			Select Case CurrentQuery.State
				Case 1	'Browsing - não precisa editar registro do contrato nem controlar transação
					vHandleContrato = CurrentQuery.FieldByName("HANDLE").AsInteger

				Case 2	'Editing
					vHandleContrato = CurrentQuery.FieldByName("HANDLE").AsInteger
					If Not InTransaction Then
						StartTransaction
					End If
				Case 3	'Inserting
					vHandleContrato = 0
					If Not InTransaction Then
						StartTransaction
					End If
			End Select

			If vsXMLContainerEnderecos = "Vazio" Then
				vsXMLContainerEnderecos = ""
			End If
			If vsXMLEnderecosExcluidos = "Vazio" Then
				vsXMLEnderecosExcluidos = ""
			End If

			Dim vResultado As Long
			vResultado = dllBSInterface0028.Contrato( CurrentSystem, vHandleContrato, _
														viHEnderecoCorrespondencia, vsXMLContainerEnderecos, _
														vsXMLEnderecosExcluidos, vsMensagem)

			If vResultado = 1 Then
				vsXMLContainerEnderecos = ""
				vsXMLEnderecosExcluidos = ""
				Err.Raise(1, Err, vsMensagem)
			Else
				Select Case CurrentQuery.State
					Case 1
						'Contrato não está em edição
						If (viHEnderecoCorrespondencia <> CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger) Then
							'só altera se houver troca de registro de endereço
							GravaEndereco( CurrentQuery.FieldByName("HANDLE").AsInteger, viHEnderecoCorrespondencia)
						End If
						vsXMLContainerEnderecos = ""
						vsXMLEnderecosExcluidos = ""
					Case 2, 3
						'Efetua preenchimento com os novos valores
						If viHEnderecoCorrespondencia > 0 Then
							CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger = viHEnderecoCorrespondencia
						Else
							CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").Clear
						End If
				End Select
			End If

			AtualizaRotulosEndereco( viHEnderecoCorrespondencia)
			Set dllBSInterface0028 = Nothing
			Exit Sub
		Except:
			msgErro = Err.Description
			numErro = Err.Number
			Set dllBSInterface0028 = Nothing
			UpdateLastUpdate("SAM_CONTRATO")
			AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )
			bsShowMessage("Falha no Cadastro de Endereço: "+Chr(13) +"(" + CStr( numErro) +")"+ msgErro, "E")
	End If
End Sub

Public Sub BOTAOFINANCEIRO_OnClick()

  If CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1 And _
                              CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "C" Then

    Dim SQL As Object
    Dim Interface As Object

    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT HANDLE FROM SFN_CONTAFIN WHERE PESSOA=" + CurrentQuery.FieldByName("PESSOA").AsString)
    SQL.Active = True
    If Not SQL.EOF Then
      Set Interface = CreateBennerObject("SamContaFinanceira.Consulta")
      Interface.Exec(CurrentSystem, SQL.FieldByName("HANDLE").AsInteger)
      Set Interface = Nothing
    Else
      bsShowMessage("Conta financeira não encontrada", "I")
    End If
    SQL.Active = False

    Set SQL = Nothing
  Else
    bsShowMessage("Somente para contrato empresarial com faturamento no contrato!", "I")

  End If
End Sub

Public Sub BOTAOREATIVAR_OnClick()
'
  If CurrentQuery.State <>1 Then
    bsShowMessage("O registro não pode estar em edição", "I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
   	bsShowMessage("O Contrato não está cancelado !", "I")
   	Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    Dim vsMensagemErro As String
    Dim viRetorno As Integer
    Dim vvContainer As CSDContainer
	Dim Interface As Object
    Set vvContainer = NewContainer

    Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


   	viRetorno = Interface.Exec(CurrentSystem, _
    								1, _
                                    "TV_FORM0009", _
           	                        "Reativação de Contrato", _
               	                    0, _
                   	                120, _
                       	            230, _
                           	        False, _
                               	    vsMensagemErro, _
                                   	vvContainer)

   	Select Case viRetorno
      	Case -1
   			bsShowMessage("Operação cancelada pelo usuário!", "I")
  		Case  1
   	  		bsShowMessage(vsMensagemErro, "I")
	End Select
  End If

  Set Interface = Nothing
  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub BOTAOVERIFICARSSO_OnClick()
  Dim BSMED001 As Object
  Set BSMED001 = CreateBennerObject("BSMED001.IncluiPacientePCMSO")

  BSMED001.Exec(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set BSMED001 = Nothing
End Sub

Public Sub CEP_OnPopup(ShowPopup As Boolean)
  ' Joldemar Moreira 12/06/2003
  ' SMS 16059
  Dim vHandle As String
  Dim Interface As Object
  ShowPopup = False
  Set Interface = CreateBennerObject("ProcuraCEP.Rotinas")
  Interface.Exec(CurrentSystem, vHandle)

  If vHandle <>"" Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Add("SELECT ESTADO    , ")
    SQL.Add("       MUNICIPIO , ")
    SQL.Add("       BAIRRO    , ")
    SQL.Add("       LOGRADOURO, ")
    SQL.Add("       COMPLEMENTO ")
    SQL.Add("  FROM LOGRADOUROS ")
    SQL.Add("  WHERE CEP = :Handle ")
    SQL.Add("UNION ")
    SQL.Add("  SELECT E.HANDLE ESTADO, ")
    SQL.Add("         M.HANDLE CIDADE, ")
    SQL.Add("         NULL, ")
    SQL.Add("         NULL, ")
    SQL.Add("         NULL  ")
    SQL.Add("    FROM MUNICIPIOS  M ")
    SQL.Add("    JOIN ESTADOS E ON (M.ESTADO    = E.HANDLE ) ")
    SQL.Add("   WHERE M.CEP =  :Handle ")
    SQL.ParamByName("HANDLE").Value = vHandle
    SQL.Active = True


    CurrentQuery.Edit
    CurrentQuery.FieldByName("CEP").Value = vHandle 'SQL.FieldByName("CEP").AsString
    CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
    CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
    CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
    CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
    CurrentQuery.FieldByName("COMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString

  End If

  Set Interface = Nothing

End Sub

Public Sub DATAADESAO_OnExit()
  If CurrentQuery.State = 3 Or(CurrentQuery.State = 2 And CurrentQuery.FieldByName("DATAADESAO").AsDateTime <>VDataUltimoReajuste)Then
    CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").Value = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
  End If
End Sub

Public Sub DIGITAR_OnClick()
  Dim BSInterface0011 As Object

  Set BSInterface0011 = CreateBennerObject("BSINTERFACE0011.DigitarBeneficiario")

  BSInterface0011.Exec(CurrentSystem, _
	                   0, _
					   CurrentQuery.FieldByName("HANDLE").AsInteger, _
					   0)

  Set BSInterface0011 = Nothing
End Sub

Public Sub GRUPOEMPRESARIAL_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|SAM_GRUPOCONTRATO.DESCRICAO|SAM_CONVENIO.DESCRICAO|DATAADESAO"

  If CurrentQuery.State = 2 Then
    vCriterio = "SAM_CONTRATO.HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString + "  AND  "
  End If
  vCriterio = vCriterio + "SAM_CONTRATO.DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND SAM_CONTRATO.NAOINCLUIRBENEFICIARIO = 'N' "
  vCriterio = vCriterio + "AND SAM_CONTRATO.TABTIPOCONTRATO = 1 "
  vCriterio = vCriterio + "AND SAM_CONTRATO.GRUPOEMPRESARIAL IS  NULL "
  vCriterio = vCriterio + "AND SAM_CONTRATO.CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString

  vCampos = "Nº do Contrato|Contratante|Grupo Contrato|Convênio|Data Adesão"

  vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO|SAM_GRUPOCONTRATO[SAM_CONTRATO.GRUPOCONTRATO = SAM_GRUPOCONTRATO.HANDLE]|SAM_CONVENIO[SAM_CONTRATO.CONVENIO = SAM_CONVENIO.HANDLE]", vColunas, 2, vCampos, vCriterio, "Contratos", True, "")

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("GRUPOEMPRESARIAL").Value = vHandle
  End If

  Set Interface = Nothing
End Sub



Public Sub PESSOA_OnAfterSearch()
' VERIFICA A CONTA FINANCEIRA
  If CurrentQuery.FieldByName("PESSOA").AsInteger > 0 Then
    Dim Erro As Long
    Dim InterfaceFin As Object
    Set InterfaceFin = CreateBennerObject("FINANCEIRO.ContaFin")
    Erro = InterfaceFin.Cadastro(CurrentSystem, CurrentQuery.FieldByName("PESSOA").AsInteger, 3, 0)
    If Erro <= 0 Then
      MsgBox "Erro " + Str(Erro) + " ao criar Conta Financeira"
    End If
    Set InterfaceFin = Nothing
  End If
  'FIM VERIFICA A CONTA FINANCEIRA
End Sub

'Public Sub PESSOA_OnChange()
  'If CurrentQuery.State =3 And  _
  '   CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger =1 Then
  '   Dim SQL As Object
  '   Set SQL =NewQuery

  '   SQL.Add("SELECT CEP, ESTADO, MUNICIPIO, BAIRRO, LOGRADOURO, NUMERO, COMPLEMENTO")
  '   SQL.Add("FROM SFN_PESSOA")
  '   SQL.Add("WHERE HANDLE = :HPESSOA")
  '   SQL.ParamByName("HPESSOA").Value =CurrentQuery.FieldByName("PESSOA").AsInteger
  '   SQL.Active =True

  '   If Not SQL.EOF Then
  ' CurrentQuery.FieldByName("CEP").Value =SQL.FieldByName("CEP").AsString
  ' CurrentQuery.FieldByName("ESTADO").Value =SQL.FieldByName("ESTADO").AsInteger
  ' CurrentQuery.FieldByName("MUNICIPIO").Value =SQL.FieldByName("MUNICIPIO").AsInteger
  ' CurrentQuery.FieldByName("BAIRRO").Value =SQL.FieldByName("BAIRRO").AsString
  ' CurrentQuery.FieldByName("LOGRADOURO").Value =SQL.FieldByName("LOGRADOURO").AsString
  ' CurrentQuery.FieldByName("NUMERO").Value =SQL.FieldByName("NUMERO").AsInteger
  ' CurrentQuery.FieldByName("COMPLEMENTO").Value =SQL.FieldByName("COMPLEMENTO").AsString
  '   End If

  '   Set SQL =Nothing
  'End If
'End Sub

Public Sub PLANO_OnExit()
  Dim SQL2 As Object
  Set SQL2 = NewQuery

  SQL2.Add("SELECT CONTABPAG, CONTABREC FROM SAM_PLANO WHERE HANDLE =:PLANO")
  SQL2.ParamByName("PLANO").Value = CurrentQuery.FieldByName("PLANO").AsInteger
  SQL2.Active = True

  If CurrentQuery.State = 3 Then
    CurrentQuery.FieldByName("CONTABPAG").Value = SQL2.FieldByName("CONTABPAG").AsInteger
    CurrentQuery.FieldByName("CONTABREC").Value = SQL2.FieldByName("CONTABREC").AsInteger
  End If

  Set SQL2 = Nothing
End Sub

Public Sub PLANO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  'Daniela -SMS 12220 -Convênio no registro da ANS
  vCriterio = "CONVENIO = " + CurrentQuery.FieldByName("CONVENIO").AsString

  vColunas = "DESCRICAO|DATACRIACAO|DATAVALIDADE"

  vCampos = "Descrição do Plano|Data de Criação|Data de Validade"

  vHandle = Interface.Exec(CurrentSystem, "SAM_PLANO", vColunas, 1, vCampos, vCriterio, "Plano", True, "")

  If vHandle <>0 Then
    '    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANO").Value = vHandle
  End If

  Set Interface = Nothing

  'Daniela -Foi tirada do PLANO_OnChange,porque não estava funcionando...
  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT * FROM SAM_PLANO WHERE HANDLE = :PLANO")
    SQL.ParamByName("PLANO").Value = CurrentQuery.FieldByName("PLANO").AsInteger
    SQL.Active = True
    CurrentQuery.FieldByName("DIASRESTRICAOFINANCEIRA").Value = SQL.FieldByName("DIASRESTRICAOFINANCEIRA").AsInteger
    CurrentQuery.FieldByName("DIASATRASO").Value = SQL.FieldByName("DIASPERMITIDOSINADIMPLENCIA").AsInteger
    CurrentQuery.FieldByName("BONIFICAMAJORAATENDIMENTO").Value = SQL.FieldByName("BONIFICAMAJORAATENDIMENTO").AsString
    Set SQL = Nothing
  End If

End Sub

'Public Sub PROXIMOVENCIMENTO_OnExit()
  'sms 24767 fernando
  'tem que estar em estado de inserção,pois so pode receber valor em estado de edição/alteração
  'If CurrentQuery.State =3 Then
  '   If Not CurrentQuery.FieldByName("PROXIMOVENCIMENTO").IsNull Then
  '      MsgBox("Não pode ser preenchido quando esta cadastrando o contrato, somente em estado de edição!")
  '      CurrentQuery.FieldByName("PROXIMOVENCIMENTO").Clear
  '      Exit Sub
  '   End If
  'End If
'End Sub

Public Sub RELATORIOAVISOSUSPENSAO_OnBtnClick()
  'Juliano 30/11/2001 ------------------------------------------------------------------------------------------------------------
  Dim OLEAutorizador As Object
  Dim handlexx As Long
  On Error GoTo cancel
  Set OLEAutorizador = CreateBennerObject("Procura.Procurar")
  handlexx = OLEAutorizador.Exec(CurrentSystem, "R_RELATORIOS", "NOME|CODIGO", 1, "Relatório", "CODIGO = 'BEN004B'", "Procura por Relatórios", True, "")
  If handlexx <>0 Then
    Dim SQL As Object
    Set SQL = NewQuery
    SQL.Add("SELECT CODIGO FROM R_RELATORIOS WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").Value = handlexx
    SQL.Active = True
    If CurrentQuery.State = 1 Then
      CurrentQuery.Edit
    End If
    CurrentQuery.FieldByName("RELATORIOAVISOSUSPENSAO").Value = SQL.FieldByName("CODIGO").AsString
  End If
  Set OLEAutorizador = Nothing
cancel :
  '--------------------------------------------------------------------------------------------------------------------------------
End Sub

Public Sub TABLE_AfterCancel()
	If InTransaction Then
		Rollback
		vsXMLContainerEnderecos = ""
		vsXMLEnderecosExcluidos = ""
	End If
End Sub

Public Sub TABLE_AfterCommitted()
	If InTransaction Then
		Commit
	End If
End Sub

Public Sub TABLE_AfterInsert()
  If(CurrentQuery.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2)And(Not CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").IsNull)Then
	If WebMode Then
		MOTIVOBLOQUEIO.WebLocalWhere = "HANDLE NOT IN (@CAMPO(MOTIVOBLOQUEIOAUTOMATICO))"
	ElseIf VisibleMode Then
  		MOTIVOBLOQUEIO.LocalWhere = "HANDLE NOT IN (@MOTIVOBLOQUEIOAUTOMATICO)"
  	End If
  End If


  If(Not CurrentQuery.FieldByName("MOTIVOBLOQUEIO").IsNull)Then
	If WebMode Then
		MOTIVOBLOQUEIOAUTOMATICO.WebLocalWhere = "HANDLE NOT IN (@CAMPO(MOTIVOBLOQUEIO))"
	ElseIf VisibleMode Then
	  	MOTIVOBLOQUEIOAUTOMATICO.LocalWhere = "HANDLE NOT IN (@MOTIVOBLOQUEIO)"
	End If
  End If

  HerdaParametro

End Sub

Public Sub TABLE_AfterPost()

  '	Inserir_origens_carencia_do_plano

  vNumeroFamiliaAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString
  vNumeroBenefAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString

  If EstadoTabela = 3 Then
    On Error GoTo Erro
    Dim Interface As Object
    Set Interface = CreateBennerObject("CONTRATO.ContratoInterface")
    If Interface.Inclui(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger) = 0 Then ' 0=fazer rollback  1=Completar Transação
      If bsShowMessage("Processo de duplicação para o contrato não pode ser completado !" + Chr(13) + _
                "                    Cancelar a inclusão do contrato?", "Q") = 6 Then 'usuário optou por cancelar a inclusão
        Err.Raise(1,"",Err.Description ) 'sms 59930 - a pedido do larini
        Exit Sub
      End If
    End If
    Set Interface = Nothing
  End If

  If EstadoTabela = 3 Then
    If GeraRegContratoCartaoMotivo Then
      bsShowMessage("Tabela Motivo de Emissão do Cartão com problemas - " + Chr(13) + _
             "Não foram gerados os motivos de emissão de cartão para o Contrato", "I")

    Else
      Dim SQL As Object
      Set SQL = NewQuery
      SQL.Add("SELECT HANDLE FROM SAM_CARTAOMOTIVO")
      SQL.Add(" WHERE TIPOMOTIVO <> 'C'")
      SQL.Active = True
      If SQL.EOF Then
        bsShowMessage("Tabela Motivo de Cartão com problemas - " + Chr(13) + _
               "Não foram gerados os motivos de emissão de cartão para o Contrato", "I")

      Else

        While Not SQL.EOF
          Dim QSU As Object
          Set QSU = NewQuery
          QSU.Add("SELECT HANDLE FROM SAM_CONTRATO_CARTAOMOTIVO WHERE CONTRATO=:CONTRATO AND CARTAOMOTIVO = :CARTAOMOTIVO")
          QSU.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
          QSU.ParamByName("CARTAOMOTIVO").AsInteger = SQL.FieldByName("HANDLE").AsInteger
          QSU.Active = True

          If QSU.EOF Then

            Dim ISQL As Object
            Set ISQL = NewQuery
            ISQL.Add("INSERT INTO SAM_CONTRATO_CARTAOMOTIVO ( HANDLE,  CONTRATO,  CARTAOMOTIVO,  COBRAREMISSAO,  TAXACARTAO)")
            ISQL.Add("                               VALUES (:HANDLE, :CONTRATO, :CARTAOMOTIVO, :COBRAREMISSAO, :TAXACARTAO)")

            ISQL.ParamByName("HANDLE").Value = NewHandle("SAM_CONTRATO_CARTAOMOTIVO")
            ISQL.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("handle").AsInteger
            ISQL.ParamByName("CARTAOMOTIVO").Value = SQL.FieldByName("handle").AsInteger
            ISQL.ParamByName("COBRAREMISSAO").Value = "N"
            ISQL.ParamByName("TAXACARTAO").Value = 0

            ISQL.ExecSQL

          End If

          SQL.Next

        Wend

      End If
    End If
  End If


  If EstadoTabela = 3 Then
    Dim SQL2 As Object
    Set SQL2 = NewQuery
    SQL2.Add("SELECT HANDLE, evento, EXIGEAUTORIZACAO FROM SAM_PLANO_EVENTOS WHERE PLANO = :PLANO")
    SQL2.ParamByName("PLANO").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
    SQL2.Active = True


    While Not SQL2.EOF
      Dim ISQL2 As Object
      Set ISQL2 = NewQuery
      ISQL2.Add("INSERT INTO SAM_CONTRATO_EVENTOS ( HANDLE,  EVENTO, CONTRATO, EXIGEAUTORIZACAO)")
      ISQL2.Add("                          VALUES   (:HANDLE, :EVENTO, :CONTRATO, :EXIGEAUTORIZACAO)")

      ISQL2.ParamByName("HANDLE").Value = NewHandle("SAM_CONTRATO_EVENTOS")
      ISQL2.ParamByName("EVENTO").Value = SQL2.FieldByName("EVENTO").AsInteger
      ISQL2.ParamByName("CONTRATO").Value = RecordHandleOfTable("SAM_CONTRATO")
      ISQL2.ParamByName("EXIGEAUTORIZACAO").Value = SQL2.FieldByName("EXIGEAUTORIZACAO").AsString

      ISQL2.ExecSQL

      SQL2.Next

    Wend



  End If


  AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )



	HerdarRelacTpDepEEstadoCivil(CurrentQuery.FieldByName("HANDLE").AsInteger)

  'CurrentQuery.Active =False
  'CurrentQuery.Active =True

  'Inserir_origens_carencia_do_plano

  If vsModoEdicao = "A" Then
    Dim vsMensagem  As String
    Dim viRetorno   As Integer
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

  If VisibleMode Then
    UpdateLastUpdate("SAM_ENDERECO")
    SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
  End If

  Exit Sub

ERRO :
  bsShowMessage("Processo de duplicação para o contrato não pode ser completado !" + Chr(13) + Str(Error), "E")
  Err.Raise(1,"",Err.Description ) 'sms 59930 - a pedido do larini

End Sub

'SHIBA 11/2001

Public Sub TABLE_AfterScroll()
  If CurrentQuery.State <> 1 Then
    BOTAOCANCELAR.Visible          = False
    BOTAOCARTAO.Visible            = False
    BOTAOFINANCEIRO.Visible        = False
    BOTAOREATIVAR.Visible          = False
    BOTAOADICIONARPLANO.Visible    = False
    BOTAOALTERAPADRAOPRECO.Visible = False
    BOTAOBENEFATIVOS.Visible       = False
    BOTAOCONTRATO.Visible          = False
    BOTAODIACOBRANCA.Visible       = False
    BOTAOVERIFICARSSO.Visible      = False
    'BOTAOENDERECO.Visible          = True
  Else
    BOTAOCANCELAR.Visible          = True
    BOTAOCARTAO.Visible            = True
    BOTAOFINANCEIRO.Visible        = True
    BOTAOREATIVAR.Visible          = True
    BOTAOADICIONARPLANO.Visible    = True
    BOTAOALTERAPADRAOPRECO.Visible = True
    BOTAOBENEFATIVOS.Visible       = True
    BOTAOCONTRATO.Visible          = True
    BOTAODIACOBRANCA.Visible       = True
    BOTAOVERIFICARSSO.Visible      = True
    'BOTAOENDERECO.Visible          = False
  End If

  NUMEROBENEFAUTOMATICO.Visible = True

  If CurrentQuery.State = 3 Then
  	PreparaNumeracaoContrato
  Else
    CONTRATO.ReadOnly = True
  End If

  ' Verifica se pode modificar o flag de numero de beneficiário automático, se já existir beneficiário cadastrado não pode alterar
  If (CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString = "S") Or (CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString = "") Then
    Dim qBenef As Object
    Set qBenef = NewQuery

    qBenef.Add("SELECT COUNT(*) QTDBENEF FROM SAM_BENEFICIARIO WHERE CONTRATO = :HANDLECONTRATO")
    qBenef.ParamByName("HANDLECONTRATO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    qBenef.Active = True

    If (qBenef.FieldByName("QTDBENEF").AsInteger > 0) Then
      NUMEROBENEFAUTOMATICO.ReadOnly = True
    Else
      NUMEROBENEFAUTOMATICO.ReadOnly = False
    End If
    Set qBenef = Nothing
  End If

  Dim vCondicao As String

  If WebMode Then
  	If CurrentQuery.State = 2 Then
    	vCondicao = "SAM_CONTRATO.HANDLE <> @CAMPO(HANDLE)) AND  "
  	End If

  	vCondicao = vCondicao + "SAM_CONTRATO.DATACANCELAMENTO IS NULL "
  	vCondicao = vCondicao + "AND SAM_CONTRATO.NAOINCLUIRBENEFICIARIO = 'N' "
  	vCondicao = vCondicao + "AND SAM_CONTRATO.TABTIPOCONTRATO = 1 "
  	vCondicao = vCondicao + "AND SAM_CONTRATO.GRUPOEMPRESARIAL IS  NULL "
  	vCondicao = vCondicao + "AND SAM_CONTRATO.CONVENIO = @CAMPO(CONVENIO)"

  	GRUPOEMPRESARIAL.WebLocalWhere = vCondicao


	PLANO.WebLocalWhere = "CONVENIO = @CAMPO(CONVENIO)"
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT HANDLE                                                                   ")
  SQL.Add("  FROM SAM_CONTRATO_SUSPENSAO                                                   ")
  SQL.Add(" WHERE CONTRATO = :CONTRATO                                                     ")
  SQL.Add("   AND ((DATAFINAL IS NULL AND DATAINICIAL <= :DATA) OR                         ")
  SQL.Add("        (DATAFINAL IS NOT NULL AND DATAFINAL >= :DATA AND DATAINICIAL <= :DATA))")
  SQL.ParamByName("CONTRATO").AsInteger  = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("DATA"    ).AsDateTime = ServerDate
  SQL.Active = True

  If (Not SQL.EOF) Then
    ROTULOSUSPENSAO.Text = "Contrato suspenso!"
  Else
    ROTULOSUSPENSAO.Text = " "
  End If

  'Para contratos empresariais...
  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) Then
    'O campo 'Local de faturamento contrib. social' somente poderá ser editado
    'se não existirem faturas para o contrato ou para algum beneficiário do contrato.
    'De acordo com a SMS 60048.
    Dim qFaturas As Object
    Set qFaturas = NewQuery

    qFaturas.Clear
    qFaturas.Add("SELECT HANDLE               ")
    qFaturas.Add("  FROM SFN_FATURA           ")
    qFaturas.Add(" WHERE CONTRATO = :CONTRATO ")
    qFaturas.Add("   AND ROTINAFIN IS NOT NULL") 'Selecionar apenas faturas geradas pela rotina de faturamento.
    qFaturas.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qFaturas.Active = True

    If (qFaturas.EOF) Then
      LOCALFATURAMENTOCONTRIBSOCIAL.ReadOnly = False
    Else
      LOCALFATURAMENTOCONTRIBSOCIAL.ReadOnly = True
    End If
    Set qFaturas = Nothing

  End If
  SQL.Active = False


  UsaGrupoCooperativa(CurrentCompany)

  PreparaNumeracaoAutomatico

  vNumeroFamiliaAutomaticoAnterior = CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString
  vNumeroBenefAutomaticoAnterior   = CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString

  SQL.Clear
  SQL.Add("SELECT UTILIZACHECONSULTA      ")
  SQL.Add("  FROM SAM_PARAMETROSPROCCONTAS")
  SQL.Active = True

  If (CurrentQuery.FieldByName("PADRAOPRECOMODULO").AsString <> "C") Then
    BOTAOALTERAPADRAOPRECO.Visible = False
  Else
    BOTAOALTERAPADRAOPRECO.Visible = True
  End If
  SQL.Active = False

  If (Not CurrentQuery.FieldByName("PLANO").IsNull) Then
    SQL.Clear
    SQL.Add("SELECT GP.DESCRICAO           ")
    SQL.Add("  FROM SAM_PLANO      P,      ")
    SQL.Add("       SAM_GRUPOPLANO GP      ")
    SQL.Add("WHERE P.HANDLE  = :HANDLE     ")
    SQL.Add("  AND GP.HANDLE = P.GRUPOPLANO")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PLANO").AsInteger
    SQL.Active = True

    ROTULOGRUPOPLANO.Text = "Grupo plano: " + SQL.FieldByName("DESCRICAO").AsString

    SQL.Active = False
  Else
    ROTULOGRUPOPLANO.Text = ""
  End If
  Set SQL = Nothing

  vMotivoBloqueioAutomatico = CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger

  If (CurrentQuery.State <> 3) Then
    DIACOBRANCA.ReadOnly       = True
    PROXIMOVENCIMENTO.ReadOnly = False
  End If

  AtualizaRotulosEndereco( CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").AsInteger )

  If (CurrentQuery.FieldByName("INCLUIRNOPCMSO").AsString = "S") Then
    BOTAOVERIFICARSSO.Visible = True
  Else
    BOTAOVERIFICARSSO.Visible = False
  End If

If(Not WebMode) Then
     COBRANCADEIMPOSTOPOR.Pages(1).Visible = Not (CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "F")
 End If

End Sub

Public Sub UsaGrupoCooperativa(handle As Integer)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT TABTIPOGESTAO   ")
  SQL.Add("  FROM EMPRESAS        ")
  SQL.Add(" WHERE HANDLE = :HANDLE")
  SQL.ParamByName("HANDLE").AsInteger = handle
  SQL.Active = True

  If (SQL.FieldByName("TABTIPOGESTAO").AsInteger = 3) Then
    GRUPOCOOP.Visible = True
  Else
    GRUPOCOOP.Visible = False
  End If
  SQL.Active = False
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
    bsShowMessage("Não é permitido excluir o contrato por motivo de suspensão!", "E")
    CanContinue = False
    Exit Sub
  End If
  Set BSBen001Dll = Nothing

  '------------------------------------------------------------------------------------------------------------------------------------
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vsModoEdicao = "A"
  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""

  BOTAOCANCELAR.Visible          = False
  BOTAOCARTAO.Visible            = False
  BOTAOFINANCEIRO.Visible        = False
  BOTAOREATIVAR.Visible          = False
  BOTAOADICIONARPLANO.Visible    = False
  BOTAOALTERAPADRAOPRECO.Visible = False
  BOTAOBENEFATIVOS.Visible       = False
  BOTAOCONTRATO.Visible          = False
  BOTAODIACOBRANCA.Visible       = False
  BOTAOVERIFICARSSO.Visible      = False

  'If CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1 Then
  '  BOTAOENDERECO.Visible = True
  'Else
  '  BOTAOENDERECO.Visible = False
  'End If

  If(CurrentQuery.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2)And(Not CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").IsNull)Then
	If WebMode Then
		MOTIVOBLOQUEIO.WebLocalWhere = "HANDLE NOT IN (@CAMPO(MOTIVOBLOQUEIOAUTOMATICO))"
	ElseIf VisibleMode Then
  		MOTIVOBLOQUEIO.LocalWhere = "HANDLE NOT IN (@MOTIVOBLOQUEIOAUTOMATICO)"
  	End If
  End If


  If(Not CurrentQuery.FieldByName("MOTIVOBLOQUEIO").IsNull)Then
	If WebMode Then
		MOTIVOBLOQUEIOAUTOMATICO.WebLocalWhere = "HANDLE NOT IN (@CAMPO(MOTIVOBLOQUEIO))"
	ElseIf VisibleMode Then
	  	MOTIVOBLOQUEIOAUTOMATICO.LocalWhere = "HANDLE NOT IN (@MOTIVOBLOQUEIO)"
	End If
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
    bsShowMessage("Não é permitido editar o contrato por motivo de suspensão!", "E")
    CanContinue = False
    CurrentQuery.Cancel
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------

  vNaoRegistrarNoMS = CurrentQuery.FieldByName("NAOREGISTRARNOMS").AsString
  VDataUltimoReajuste = CurrentQuery.FieldByName("DATAADESAO").Value
  vCobrancaDeEvento = CurrentQuery.FieldByName("COBRANCADEEVENTO").AsString
  vDiaCobranca = CurrentQuery.FieldByName("DIACOBRANCA").AsInteger


  If WebMode Then
	If WebVisionCode = "V_SAM_CONTRATO_426" Then
		GRUPOCONTRATO.ReadOnly = True
	End If
  Else
	  If Not InTransaction Then
	    StartTransaction
	  End If
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  vsModoEdicao = "I"
  vsXMLContainerEnderecos = ""
  vsXMLEnderecosExcluidos = ""

  BOTAOCANCELAR.Visible          = False
  BOTAOCARTAO.Visible            = False
  BOTAOFINANCEIRO.Visible        = False
  BOTAOREATIVAR.Visible          = False
  BOTAOADICIONARPLANO.Visible    = False
  BOTAOALTERAPADRAOPRECO.Visible = False
  BOTAOBENEFATIVOS.Visible       = False
  BOTAOCONTRATO.Visible          = False
  BOTAODIACOBRANCA.Visible       = False
  BOTAOVERIFICARSSO.Visible      = False
  'BOTAOENDERECO.Visible          = True

  vNaoRegistrarNoMS = "N"
  'SMS 24767 FERNANDO
  DIACOBRANCA.ReadOnly = False
  PROXIMOVENCIMENTO.ReadOnly = True

  If WebMode Then
	If WebVisionCode = "V_SAM_CONTRATO_426" Then
		GRUPOCONTRATO.ReadOnly = True
	End If
  Else
  	If Not InTransaction Then
  		StartTransaction
  	End If
  End If
End Sub

'#Uses "*VerificaEmail"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vsMensagem As String



  If CurrentQuery.State = 2 Then
    If CurrentQuery.FieldByName("DIACOBRANCA").OldValue <> CurrentQuery.FieldByName("DIACOBRANCA").AsInteger Then
      bsShowMessage("O campo Dia de Cobrança não pode ser alterado, somente o campo Proximo Vencimento!", "I")
      CurrentQuery.FieldByName("DIACOBRANCA").AsInteger = CurrentQuery.FieldByName("DIACOBRANCA").OldValue
      Exit Sub
    End If
  End If


  If CurrentQuery.State = 3 Then
    CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").Value = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
  End If

  If (CurrentQuery.State = 1) Then
    bsShowMessage("Status 1", "I")
  End If

  If (CurrentQuery.FieldByName("PERMITEFATURARPFAUTORIZACAO").AsString = "S") Then
    If ((CurrentQuery.FieldByName("VALIDADEPAGAMENTO").IsNull) Or (CurrentQuery.FieldByName("VALIDADEAUTORIZACAO").IsNull)) Then
      bsShowMessage("Se permitir faturar PF na autorização deve informar as validades de pagamento e autorização", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  'Contador geral(universal) ou por empresa para o contrato.
  If ((CONTRATO.ReadOnly) And (CurrentQuery.State = 3)) Then
    CurrentQuery.FieldByName("CONTRATO").AsInteger = CriaContadorContrato
  End If

  If (CurrentQuery.FieldByName("CONTRATO").IsNull) Then
    bsShowMessage("O campo 'Contrato' é de preenchimento obrigatório.", "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = CheckNumeroAutomatico
  If (Not CanContinue) Then Exit Sub

  'Vericar se o contrato é único conforme parametros gerais.
  If (Not NumeroContratoUnico) Then
    bsShowMessage("Contrato já cadastrado.", "E")
    CanContinue = False
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  If (Not CurrentQuery.FieldByName("GRUPOEMPRESARIAL").IsNull) Then
    SQL.Clear
    SQL.Add("SELECT GRUPOEMPRESARIAL")
    SQL.Add("  FROM SAM_CONTRATO    ")
    SQL.Add(" WHERE HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("GRUPOEMPRESARIAL").AsInteger
    SQL.Active = True

    If (Not SQL.FieldByName("GRUPOEMPRESARIAL").IsNull) Then
      bsShowMessage("Grupo empresarial inválido. O contrato principal do grupo empresarial não pode ser dependente de outro contrato.", "E")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If

    SQL.Clear
    SQL.Add("SELECT HANDLE                              ")
    SQL.Add("  FROM SAM_CONTRATO                        ")
    SQL.Add(" WHERE GRUPOEMPRESARIAL = :GRUPOEMPRESARIAL")
    SQL.ParamByName("GRUPOEMPRESARIAL").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True

    If (Not SQL.EOF) Then
      bsShowMessage("Um contrato principal de grupo empresarial não pode ser dependente de outro contrato.", "Ë")
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If
  Set SQL = Nothing

  If (CurrentQuery.FieldByName("TIPODOCUMENTODESTINOBENEF").IsNull) Then
    Dim qBenef As Object
    Set qBenef = NewQuery

    qBenef.Clear
    qBenef.Add("SELECT COUNT(1) QTDE                       ")
    qBenef.Add("  FROM SAM_BENEFICIARIO                    ")
    qBenef.Add(" WHERE CONTRATO = :CONTRATO                ")
    qBenef.Add("   AND (DESTINOCOBCONTRIBUICAOSOCIAL = 2 OR")
    qBenef.Add("        DESTINOCOBMENSALIDADE        = 2 OR")
    qBenef.Add("        DESTINOCOBPFSERVICO          = 2 OR")
    qBenef.Add("        DESTINOCOBFRQURGEMERG        = 2)  ")
    qBenef.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qBenef.Active = True

    If (qBenef.FieldByName("QTDE").AsInteger > 0) Then
      bsShowMessage("O parâmetro 'Tipo documento para cobrança do beneficiário' deve ser informado. Existem beneficiários configurados como recebedores de cobrança.", "E")
      CanContinue = False
      Set qBenef = Nothing
      Exit Sub
    End If
    Set qBenef = Nothing
  End If

  Dim vDep      As Long
  Dim vFam      As Long
  Dim vCont     As Long
  Dim vEmp      As Long
  Dim Interface As Object

  Set Interface = CreateBennerObject("SAMBENEFICIARIO.Cadastro")
  Interface.ContaDigitosComposicaoBenef(CurrentSystem, vDep, vFam, vCont, vEmp)
  Set Interface = Nothing

  If ((vCont > 0) And (Len(CurrentQuery.FieldByName("CONTRATO").AsString) > vCont)) Then
    bsShowMessage("Contrato com mais dígitos do que definido na composição do código do beneficiário.", "E")
    CanContinue = False
    Exit Sub
  End If

  If (Not CurrentQuery.FieldByName("EMAIL").IsNull) Then
    If (Not VerificaEmail(CurrentQuery.FieldByName("EMAIL").AsString)) Then
      bsShowMessage("E-mail inválido.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If (Not CurrentQuery.FieldByName("EMAILRESPONSAVEL").IsNull) Then
    If (Not VerificaEmail(CurrentQuery.FieldByName("EMAILRESPONSAVEL").AsString)) Then
      bsShowMessage("E-mail inválido.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  EstadoTabela = CurrentQuery.State

  If (CurrentQuery.FieldByName("MOTIVOBLOQUEIO").IsNull) Then
    CurrentQuery.FieldByName("DATABLOQUEIO").Clear
  Else
    CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime = ServerDate
  End If

  If ((CurrentQuery.FieldByName("FRANQUIA").AsString = "S") And (CurrentQuery.FieldByName("TIPOCONTAGEMFRANQUIA").IsNull)) Then
    bsShowMessage("Para contratos com franquia deve-se informar o tipo de contagem da mesma.", "E")
    CanContinue = False
  End If

  If (CurrentQuery.State <> 3) Then
    'Verificar se foi alterada a forma de pagamento.
    If (CurrentQuery.FieldByName("COBRANCADEEVENTO").AsString <> vCobrancaDeEvento) Then
     bsShowMessage("Alterado tipo de cobrança de eventos de" + CurrentQuery.FieldByName("COBRANCADEEVENTO").AsString + "para" +  vCobrancaDeEvento + "!", "I")
    End If

   'Verificar se foi alterado o dia de cobrança.
    If (CurrentQuery.FieldByName("DIACOBRANCA").AsInteger <> vDiaCobranca) Then
		bsShowMessage("Alterado dia de cobrança de eventos de" + CStr(CurrentQuery.FieldByName("DIACOBRANCA").AsInteger) + "para" +  CStr(vDiaCobranca) + "!", "I")
    End If
  End If

  If (vNaoRegistrarNoMS <> CurrentQuery.FieldByName("NaoRegistrarNoMS").AsString) Then
    Dim Texto As String
    If (CurrentQuery.FieldByName("NaoRegistrarNoMS").AsString = "S") Then
      Texto = "O Contrato foi modificado para NÃO registrar os beneficiários no Ministério da Saúde."
    Else
      Texto = "O Contrato foi modificado para registrar os beneficiários no Ministério da Saúde."
    End If
    Texto = Texto + " Confirma?"
    If (MsgBox(Texto, vbYesNo, "contrato") <> vbYes) Then
      CanContinue = False
      Exit Sub
    End If
  End If

  'TABTIPOCONTRATO:
  '  1 = EMPRESARIAL;
  '  2 = FAMILIAR;
  '  3 = INDIVIDUAL.
  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger <> 1) Then
    If (CurrentQuery.State = 3) Then
      CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "F"
    End If
    CurrentQuery.FieldByName("DIACOBRANCA").Clear
    CurrentQuery.FieldByName("COBRANCANOMESSEGUINTE").AsString = "N"
  End If

  If ((CurrentQuery.FieldByName("TABFOLHAPAGAMENTO").AsInteger = 2) And _
      (CurrentQuery.FieldByName("TABPERMITERECEBERMENSALOUTROS").AsInteger = 2) And _
      (CurrentQuery.FieldByName("CODIGOFOLHAMENSALIDADE").IsNull)) Then
    bsShowMessage("O campo 'Código folha para mensalidade' é de preenchimento obrigatório.", "E")
    CanContinue = False
    Exit Sub
  End If

  If ((CurrentQuery.FieldByName("TABFOLHAPAGAMENTO").AsInteger = 1) And _
      (CurrentQuery.FieldByName("TABENVIARMENSALFOLHATITULAR").AsInteger = 2) And _
      (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) And _
      (CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString <> "F")) Then
    bsShowMessage("A opção de enviar a mensalidade para a folha do titular é permitida apenas para contratos com faturamento na 'Família'.", "E")
    CanContinue = False
    Exit Sub
  End If

  'Checar o dia de cobrança.
  Dim SQLTipoFat As Object
  Set SQLTipoFat = NewQuery

  SQLTipoFat.Clear
  SQLTipoFat.Add("SELECT CODIGO             ")
  SQLTipoFat.Add("  FROM SIS_TIPOFATURAMENTO")
  SQLTipoFat.Add(" WHERE HANDLE = :HANDLE   ")
  SQLTipoFat.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
  SQLTipoFat.Active = True

  If ((SQLTipoFat.FieldByName("CODIGO").Value <> 130) And (CurrentQuery.FieldByName("TABFOLHAPAGAMENTO").AsInteger = 2)) Then
    bsShowMessage("Apenas contratos de autogestão podem ter relacionamento com folha de pagamento.", "E")
    CanContinue = False
    Set SQLTipoFat = Nothing
    Exit Sub
  End If

  If ((SQLTipoFat.FieldByName("CODIGO").AsInteger = 130) And (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger <> 1)) Then
    bsShowMessage("Contratos de autogestão devem ser empresariais.", "E")
    CanContinue = False
    Set SQLTipoFat = Nothing
    Exit Sub
  End If

  If ((SQLTipoFat.FieldByName("CODIGO").Value = 130) Or (CurrentQuery.FieldByName("LOCALFATURAMENTO").Value = "C")) Then
    If (CurrentQuery.FieldByName("DIACOBRANCA").IsNull) Then
      bsShowMessage("Para faturamento no contrato ou autogestão deve-se informar o dia de cobrança.", "E")
      CanContinue = False
      Set SQLTipoFat = Nothing
      Exit Sub
    Else
      'Alimentar DiaCobrancaOriginal na inclusão e se contrato ainda não foi faturado.
      If (CurrentQuery.State = 3) Then 'Inclusão.
        CurrentQuery.FieldByName("DIACOBRANCAORIGINAL").AsInteger = CurrentQuery.FieldByName("DIACOBRANCA").AsInteger
      Else
        'Se o contrato ainda não foi faturado alimenta DiaCobrancaOriginal.
        Dim SQLFat As Object
        Set SQLFat = NewQuery

        SQLFat.Clear
        SQLFat.Add("SELECT A.HANDLE                            ")
        SQLFat.Add("  FROM SFN_CONTAFIN A,                     ")
        SQLFat.Add("       SFN_FATURA   B                      ")
        SQLFat.Add(" WHERE A.PESSOA          = :PESSOA         ")
        SQLFat.Add("   AND B.TIPOFATURAMENTO = :TIPOFATURAMENTO")
        SQLFat.Add("   And B.CONTAFINANCEIRA = A.HANDLE        ")
        SQLFat.ParamByName("TIPOFATURAMENTO").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
        SQLFat.ParamByName("PESSOA"         ).AsInteger = CurrentQuery.FieldByName("PESSOA").AsInteger
        SQLFat.Active = True

        If (SQLFat.EOF) Then
          If CurrentQuery.FieldByName("DIACOBRANCA").AsInteger > 0 Then
            CurrentQuery.FieldByName("DIACOBRANCAORIGINAL").AsInteger = CurrentQuery.FieldByName("DIACOBRANCA").AsInteger
          End If
        End If
        SQLFat.Active = False
        Set SQLFat = Nothing
      End If
    End If
  Else
    If (Not CurrentQuery.FieldByName("DIACOBRANCA").IsNull) Then
      bsShowMessage("Deve-se informar o dia de cobranca somente para faturamento no contrato ou autogestão.", "E")
      CanContinue = False
      Set SQLTipoFat = Nothing
      Exit Sub
    End If
  End If
  Set SQLTipoFat = Nothing

  CanContinue = CheckVigenciaPlano
  If (Not CanContinue) Then Exit Sub

  'Checar CNPJ se o contrato é empresarial.
  'TABTIPOCONTRATO:
  '  1 = EMPRESARIAL;
  '  2 = FAMILIAR;
  '  3 = INDIVIDUAL.

  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) And (CurrentQuery.FieldByName("PESSOA").AsInteger = 0) Then
	bsShowMessage("O campo Pessoa é obrigatório quando o tipo de contrato for empresarial.", "E")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) Then
    If (CurrentQuery.FieldByName("CNPJ").IsNull) Then
      If (CurrentQuery.FieldByName("CPFRESPONSAVEL").IsNull) Then
        bsShowMessage("CNPJ ou CPF do responsável deve ser informado.", "E")
        CanContinue = False
        Exit Sub
      Else
        If (Not IsValidCPF(CurrentQuery.FieldByName("CPFRESPONSAVEL").AsString)) Then
          bsShowMessage("O CPF do responsável é inválido.", "E")
          CanContinue = False
          Exit Sub
        End If
      End If
    Else
      If Not IsValidCGC(CurrentQuery.FieldByName("CNPJ").AsString)Then
        bsShowMessage("Contrato empresarial - CNPJ inválido.", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

    If (CurrentQuery.FieldByName("RAZAOSOCIAL").IsNull) Then
      bsShowMessage("O campo 'Razão social' é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DATAABERTURA").IsNull) Then
      bsShowMessage("O campo referente à data de abertura é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("NOMERESPONSAVEL").IsNull) Then
      bsShowMessage("O campo 'Nome' do responsável é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If

  End If

  'Verificar o padrão de correspondência.
  'TABTIPOCONTRATO:
  '  1 = EMPRESARIAL;
  '  2 = FAMILIAR;
  '  3 = INDIVIDUAL.
  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger <> 1) Then
    If (CurrentQuery.FieldByName("INFORMATIVOS").AsString = "C") Then
      bsShowMessage("Contratos Familiar/Individual não podem receber informativos no contrato.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("CARTAO").AsString = "C") Then
      bsShowMessage("Contratos Familiar/Individual não podem receber cartões no contrato.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("COBRANCA").AsString = "C") Then
      bsShowMessage("Contratos Familiar/Individual não podem receber cobranças no contrato.", "E")
      CanContinue = False
      Exit Sub
    End If

    'Verificar o local de faturamento.
    If (CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "C") Then
      bsShowMessage("Contratos familiar/individual não podem ter faturamento no contrato.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  'Verificar se existe familia sem titular quando a correspondência está na família.
  Dim vbResultado As Boolean

  Set Interface = CreateBennerObject("SAMENDERECO.Localiza")
  vbResultado = Interface.ValidaCorrespContrato(CurrentSystem, _
                                                CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                CurrentQuery.FieldByName("COBRANCA").AsString, _
                                                CurrentQuery.FieldByName("INFORMATIVOS").AsString, _
                                                vsMensagem)
  Set Interface = Nothing

  If (Not vbResultado) And WebMode Then
    bsShowMessage(vsMensagem, "E")
  End If

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT UTILIZACHECONSULTA      ")
  SQL.Add("  FROM SAM_PARAMETROSPROCCONTAS")
  SQL.Active = True

  If ((CurrentQuery.FieldByName("CHECONSULTAPORBENEF").IsNull) And (SQL.FieldByName("UTILIZACHECONSULTA").AsBoolean)) Then
    bsShowMessage("O campo 'Núm. de Cheque Consulta por Beneficiário' é de preenchimento obrigatório.", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If
  Set SQL = Nothing

  UsaGrupoCooperativa(CurrentCompany)

  If (GRUPOCOOP.Visible) Then
    If (CurrentQuery.FieldByName("LOCATEND").IsNull) Then
      bsShowMessage("O campo 'Local atend.' (cooperativa) é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("LOCCOB").IsNull) Then
      bsShowMessage("O campo 'Local cobrança' (cooperativa) é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim SqlDataFechamento
  Set SqlDataFechamento = NewQuery

  SqlDataFechamento.Clear
  SqlDataFechamento.Add("SELECT DATAFECHAMENTO            ")
  SqlDataFechamento.Add("  FROM SAM_PARAMETROSBENEFICIARIO")
  SqlDataFechamento.Active = True

  If (CurrentQuery.State = 3) Then
    If (SqlDataFechamento.FieldByName("DATAFECHAMENTO").AsDateTime > CurrentQuery.FieldByName("DATAADESAO").AsDateTime) Then
      bsshowMessage("Não é possível cadastrar data de adesão inferior à data de fechamento - Parâmetros Gerais.", "E")
      CanContinue = False
      Set SqlDataFechamento = Nothing
      Exit Sub
    End If
  End If

  If (Not CurrentQuery.FieldByName("DATABLOQUEIO").IsNull) Then
    If (SqlDataFechamento.FieldByName("DATAFECHAMENTO").AsDateTime > CurrentQuery.FieldByName("DATABLOQUEIO").AsDateTime) Then
      bsShowMessage("Não é possível cadastrar data de bloqueio inferior à data de fechamento - Parâmetros Gerais.", "E")
      CanContinue = False
      Set SqlDataFechamento = Nothing
      Exit Sub
    End If
  End If
  Set SqlDataFechamento = Nothing

  Dim Voperadora   As Integer
  Dim SqlOperadora As Object

  Set SqlOperadora = NewQuery

  If (CurrentQuery.State = 3) Then
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT C.OPERADORA,           ")
    SQL.Add("       SO.TABTIPO,            ")
    SQL.Add("       SO.HANDLE,             ")
    SQL.Add("       SO.ADMINISTRADORA      ")
    SQL.Add("  FROM SAM_CONVENIO  C,       ")
    SQL.Add("       SAM_OPERADORA SO       ")
    SQL.Add(" WHERE C.OPERADORA = SO.HANDLE")
    SQL.Add("   AND C.HANDLE    = :CONVENIO")
    SQL.ParamByName("CONVENIO").AsInteger = CurrentQuery.FieldByName("CONVENIO").AsInteger
    SQL.Active = True

    Voperadora = SQL.FieldByName("OPERADORA").AsInteger

    SqlOperadora.Clear
    SqlOperadora.Add("SELECT MAX(DATAFINAL) DATAFINAL     ")
    SqlOperadora.Add("  FROM SAM_MSPROCESSO               ")

    If (SQL.FieldByName("TABTIPO").AsInteger = 1) Then
      SqlOperadora.Add(" WHERE OPERADORA    = :OPERADORA")
      SqlOperadora.ParamByName("OPERADORA").AsInteger = SQL.FieldByName("OPERADORA").AsInteger
    Else
      SqlOperadora.Add(" WHERE OPERADORAADM = :OPERADORA")
      SqlOperadora.ParamByName("OPERADORA").AsInteger = SQL.FieldByName("ADMINISTRADORA").AsInteger
    End If

    SqlOperadora.Add("   AND TABTIPOOPERADORA = :TABTIPO  ")
    SqlOperadora.Add("   AND SITUACAO         = 'P'       ")
    SqlOperadora.Add("   AND DATAEXPORTACAO IS NOT NULL   ")
    SqlOperadora.ParamByName("TABTIPO").AsInteger = SQL.FieldByName("tabtipo").AsInteger
    SqlOperadora.Active = True

    If (SqlOperadora.FieldByName("DATAFINAL").AsDateTime >= CurrentQuery.FieldByName("DATAADESAO").AsDateTime) Then
      bsShowMessage("A data de adesão é menor que data final do envio de beneficiários à ANS.", "E")
      CanContinue = False
      Set SQL = Nothing
      Set SqlOperadora = Nothing
      Exit Sub
    End If
    Set SQL = Nothing
    Set SqlOperadora = Nothing
  End If

  If (CurrentQuery.FieldByName("COBRARBASEPFCONTRATO").AsString = "S") Then
    If ((CurrentQuery.FieldByName("TABTIPOCONTRATO").AsString = "1") And _
        (CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger = 14) And _
        (CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "F")) Then
      CanContinue = True
    Else
      bsShowMessage("Só é permitido a marcação do campo 'Cobrar base da PF do Contrato' quando o tipo de faturamento do contrato for 'Autogestão' e o contrato for 'Empresarial' com faturamento na 'Família'.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  If (CurrentQuery.FieldByName("PERMITEPFNAIMPORTACAOBENEF").AsString = "N") Then
    Dim qPfEvento As Object
    Set qPfEvento = NewQuery

    qPfEvento.Clear
    qPfEvento.Add("SELECT COUNT(HANDLE) QTDE               ")
    qPfEvento.Add("  FROM SAM_CONTRATO_PFEVENTO            ")
    qPfEvento.Add(" WHERE CONTRATO              = :CONTRATO")
    qPfEvento.Add("   AND ATUALIZARNAIMPORTACAO = 'S'      ")
    qPfEvento.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qPfEvento.Active = True

    If (qPfEvento.FieldByName("QTDE").AsInteger > 0) Then
      bsShowMessage("Contrato possui Pf por Evento com o parâmetro 'Atualizar na importação' marcado.", "E")
      CanContinue = False
      Set qPfEvento = Nothing
      Exit Sub
    End If
    Set qPfEvento = Nothing
  End If

  If (CurrentQuery.FieldByName("TABADESAORECEBIMENTO").AsInteger = 2) Then
    If (CurrentQuery.FieldByName("PRIMEIRAPARCELANAINSCRICAO").AsString = "N") Then
      bsShowMessage("Para contratos que utilizam adesão no recebimento, o parâmetro 'Primeira parcela na rotina de inscrição' deve estar marcado.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("SEGUNDAPARCELADEPENDEPRIMEIRA").AsInteger <> 3) Then
      bsShowMessage("Para contratos que utilizam adesão no recebimento, o parâmetro 'Não faturar segunda parcela sem primeira' deve estar marcado com a opção 'Somente com a primeira paga'.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").IsNull) Then
      bsShowMessage("Informar o motivo de bloqueio automático.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (Not CurrentQuery.FieldByName("MOTIVOBLOQUEIO").IsNull) Then
      If (CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger = CurrentQuery.FieldByName("MOTIVOBLOQUEIO").AsInteger) Then
        bsShowMessage("O motivo de bloqueio e o motivo de bloqueio automático não podem ser iguais.", "E")
        CanContinue = False
        Exit Sub
      End If
    End If

    If (CurrentQuery.FieldByName("TIPODOCUMENTO").IsNull) Then
      bsShowMessage("O campo 'Tipo de documento' é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DIASVENCIMENTOPRIMEIRAPARCELA").IsNull) Then
      bsShowMessage("O campo 'Dias para vencimento primeira parcela' é de preenchimento obrigatório.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (vMotivoBloqueioAutomatico <> CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger) Then
      Dim qBloqueio As Object
      Set qBloqueio = NewQuery

      qBloqueio.Add("SELECT HANDLE                                    ")
      qBloqueio.Add("  FROM SAM_FAMILIA                               ")
      qBloqueio.Add(" WHERE CONTRATO       = :CONTRATO                ")
      qBloqueio.Add("   AND MOTIVOBLOQUEIO = :MOTIVOBLOQUEIOAUTOMATICO")
      qBloqueio.ParamByName("CONTRATO"                ).AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qBloqueio.ParamByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger = CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger
      qBloqueio.Active = True

      If (Not qBloqueio.FieldByName("HANDLE").IsNull) Then
        bsShowMessage("Motivo de bloqueio automático inválido. Existe(m) família(s) bloqueada(s) com o mesmo motivo.", "E")
        CanContinue = False
        Set qBloqueio = Nothing
        Exit Sub
      Else
        qBloqueio.Active = False

        qBloqueio.Clear
        qBloqueio.Add("SELECT HANDLE                                   ")
        qBloqueio.Add("  FROM SAM_BENEFICIARIO                         ")
        qBloqueio.Add(" WHERE CONTRATO      = :CONTRATO                ")
        qBloqueio.Add("  AND MOTIVOBLOQUEIO = :MOTIVOBLOQUEIOAUTOMATICO")
        qBloqueio.ParamByName("CONTRATO"                ).AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qBloqueio.ParamByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger = CurrentQuery.FieldByName("MOTIVOBLOQUEIOAUTOMATICO").AsInteger
        qBloqueio.Active = True

        If (Not qBloqueio.FieldByName("HANDLE").IsNull) Then
          bsShowMessage("Motivo de bloqueio automático inválido. Existe(m) beneficiário(s) bloqueado(s) com o mesmo motivo.", "E")
          CanContinue = False
          Set qBloqueio = Nothing
          Exit Sub
        End If
      End If
      Set qBloqueio = Nothing
    End If
  End If

  If (CurrentQuery.FieldByName("TABCENTROCUSTO").AsInteger = 2) Then
    Dim qBuscaCC As Object
    Set qBuscaCC = NewQuery

    qBuscaCC.Clear
    qBuscaCC.Add("SELECT COUNT(CENTROCUSTO) NUMERO,")
    qBuscaCC.Add("       FILIAL                    ")
    qBuscaCC.Add("  FROM SAM_CONTRATO_CENTROCUSTO  ")
    qBuscaCC.Add(" WHERE CONTRATO = :CONTRATO      ")
    qBuscaCC.Add("GROUP BY FILIAL                  ")
    qBuscaCC.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qBuscaCC.Active = True

    While (Not qBuscaCC.EOF)
      If (qBuscaCC.FieldByName("NUMERO").AsInteger > 1) Then
        bsShowMessage("Existe(m) filial(is) definida(s) com mais de um centro de custo. Verifique.", "E")
        CanContinue = False
        Set qBuscaCC = Nothing
        Exit Sub
      End If
      qBuscaCC.Next
    Wend
    Set qBuscaCC = Nothing
  End If

  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT CODIGO             ")
  SQL.Add("  FROM SIS_TIPOFATURAMENTO")
  SQL.Add(" WHERE HANDLE = :HANDLE   ")
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
  SQL.Active = True

  If (((SQL.FieldByName("CODIGO").AsInteger <> 110) Or _
       (CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString <> "C") Or _
       (CurrentQuery.FieldByName("COBRANCADEEVENTO").AsString = "C")) And _
      (CurrentQuery.FieldByName("PERMITEREPASSE").AsInteger = 1)) Then
    bsShowMessage("Pemite repasse somente para contratos de custo operacional, local de faturamento no contrato e cobrança de evento igual a 'Preço negociado'.", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If
  Set SQL = Nothing

  'Soares - SMS: 58872 - 22/08/2006 - Início
  'Se for marcado o flag validar prazo recibo reembolso, então se torna obrigatório o preenchimento
  'dos campos DiasPrazoModuloCobertura e DiasPrazoModuloReembolso
  If (CurrentQuery.FieldByName("VALIDARRECIBOREEMBOLSO").AsString = "S") Then
    If (CurrentQuery.FieldByName("DIASPRAZOMODULOCOBERTURA").IsNull) Then
      bsShowMessage("Campo módulo de cobertura obrigatório.", "E")
      DIASPRAZOMODULOCOBERTURA.SetFocus
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DIASPRAZOMODULOSUPLEMENTACAO").IsNull) Then
      bsShowMessage("Campo módulo de suplementação obrigatório.", "E")
      DIASPRAZOMODULOSUPLEMENTACAO.SetFocus
      CanContinue = False
      Exit Sub
    End If
  End If
  'Soares - SMS: 58872 - 22/08/2006 – Fim

  'Verificação de obrigatoriedade do endereço de correspondência
  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) And _
     (CurrentQuery.FieldByName("ENDERECOCORRESPONDENCIA").IsNull)Then
    If WebMode And _
       CurrentQuery.State = 3 Then
      bsShowMessage("Contrato Empresarial exige informações de endereço!", "I")
    Else
      CanContinue = False
      bsShowMessage("Contrato Empresarial exige informações de endereço!", "E")
    End If
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) And _
     (CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "F") And _
     (CurrentQuery.FieldByName("COBRANCADEIMPOSTOPOR").AsInteger = 2) Then
    If WebMode And _
       CurrentQuery.State = 3 Then
      bsShowMessage("Não é permitida a cobrança de imposto na pessoa responsável, quando o local de faturamento for a família!", "I")
    Else
      CanContinue = False
      bsShowMessage("Não é permitida a cobrança de imposto na pessoa responsável, quando o local de faturamento for a família!", "E")
    End If
    Exit Sub
  End If

  'Este codigo deve ser o último do evento.
  If (CurrentQuery.State = 3) Then
    If (Not CurrentQuery.FieldByName("PROXIMOVENCIMENTO").IsNull) Then
      bsShowMessage("O campo 'Próximo Vencimento' não pode ser preenchido na inclusão do contrato. Ele assumirá automaticamente a data do dia de Cobrança.", "E")
      CanContinue = False
      Exit Sub
    End If

    If (Not CurrentQuery.FieldByName("DIACOBRANCA").IsNull) Then
      CurrentQuery.FieldByName("PROXIMOVENCIMENTO").AsInteger = CurrentQuery.FieldByName("DIACOBRANCA").AsInteger
    End If
  End If
End Sub


Public Function CheckVigenciaPlano As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  CheckVigenciaPlano = True

  SQL.Add("SELECT * FROM SAM_PLANO WHERE HANDLE = :PLANO")
  SQL.ParamByName("PLANO").Value = CurrentQuery.FieldByName("PLANO").AsInteger
  SQL.Active = True

  If CurrentQuery.FieldByName("DATAADESAO").AsDateTime <SQL.FieldByName("DATACRIACAO").AsDateTime Then
    bsShowMessage("Data de adesão do Contrato inferior a criação do Plano!", "E")
    CheckVigenciaPlano = False
  Else
    If Not SQL.FieldByName("DATAVALIDADE").IsNull Then
      If CurrentQuery.FieldByName("DATAADESAO").AsDateTime >SQL.FieldByName("DATAVALIDADE").AsDateTime Then
        bsShowMessage("Data de adesão do Contrato maior que a Validade do Plano!", "E")
        CheckVigenciaPlano = False
      End If
    End If
  End If
  Set SQL = Nothing
End Function

'Public Sub TABTIPOCONTRATO_OnChange()
' If CurrentQuery.State =3 Then
'    If CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger <>1 Then    '1 empresarial 2 familiar 3 individual
'      If CurrentQuery.FieldByName("INFORMATIVOS").AsString  ="C" Then
'       CurrentQuery.FieldByName("INFORMATIVOS").Value  ="F"
'     End If
'     If CurrentQuery.FieldByName("CARTAO").AsString  ="C" Then
'       CurrentQuery.FieldByName("CARTAO").Value  ="F"
'     End If
'     If CurrentQuery.FieldByName("COBRANCA").AsString  ="C" Then
'       CurrentQuery.FieldByName("COBRANCA").Value  ="F"
'     End If
'   End If
'End If
'End Sub




Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("EMPRESA").Value = CurrentCompany

  Dim SQL As Object
  Set SQL = NewQuery


  SQL.Add("SELECT GR.TESOURARIA FROM SAM_GRUPOCONTRATO GR, SAM_CONTRATO CO")
  SQL.Add(" WHERE GR.HANDLE = CO.GRUPOCONTRATO AND GR.HANDLE = :PGRUPO")
  SQL.ParamByName("PGRUPO").Value = CurrentQuery.FieldByName("GRUPOCONTRATO").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    CurrentQuery.FieldByName("TESOURARIA").Value = SQL.FieldByName("TESOURARIA").AsInteger
  End If

  'sms 31541
  ROTULOCORRESP1.Text = ""
  ROTULOCORRESP2.Text = ""
  ROTULOCORRESP3.Text = ""
  ROTULOCORRESP4.Text = ""
  ROTULOCORRESP5.Text = ""

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOADICIONARPLANO"
			BOTAOADICIONARPLANO_OnClick
		Case "BOTAOALTERAPADRAOPRECO"
			BOTAOALTERAPADRAOPRECO_OnClick
		Case "BOTAOBENEFATIVOS"
			BOTAOBENEFATIVOS_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAODIACOBRANCA"
			BOTAODIACOBRANCA_OnClick
		Case "BOTAOFINANCEIRO"
			BOTAOFINANCEIRO_OnClick
		Case "BOTAOREATIVAR"
			BOTAOREATIVAR_OnClick
		Case "BOTAOVERIFICARSSO"
			BOTAOVERIFICARSSO_OnClick
		Case "DIGITAR"
			DIGITAR_OnClick
		Case "BOTAOCONTRATO"
			BOTAOCONTRATO_OnClick
		Case "RELATORIOAVISOSUSPENSAO"
			RELATORIOAVISOSUSPENSAO_OnBtnClick
	End Select
End Sub

Public Sub TABTIPOCONTRATO_OnChange()
  If TABTIPOCONTRATO.PageIndex = 0 Then
    CurrentQuery.FieldByName("PADRAOPRECOMODULO").Value = "C"
  Else
    CurrentQuery.FieldByName("PADRAOPRECOMODULO").Value = "F"
  End If
End Sub

Public Sub TABTIPOCONTRATO_OnChanging(AllowChange As Boolean)
  If CurrentQuery.State <>3 Then
    AllowChange = False
  End If
End Sub


Public Sub PreparaNumeracaoContrato
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT NUMEROCONTRATOAUTOMATICO FROM EMPRESAS WHERE HANDLE = :HEMPRESA")
  SQL.ParamByName("HEMPRESA").Value = CurrentCompany
  SQL.Active = True

  If SQL.FieldByName("NUMEROCONTRATOAUTOMATICO").AsString = "N" Then
    CONTRATO.ReadOnly = False
  Else
    CONTRATO.ReadOnly = True
  End If
  Set SQL = Nothing
End Sub

Public Sub PreparaNumeracaoAutomatico
  NUMEROFAMILIAAUTOMATICO.Visible = True
End Sub


Public Function NumeroContratoUnico As Boolean
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")
  SQL.Active = True
  Select Case SQL.FieldByName("NUMEROCONTRATOUNICO").AsInteger
    Case 1
      SQL.Clear
      SQL.Add("SELECT HANDLE FROM SAM_CONTRATO WHERE CONTRATO = :CONTRATO")
      If CurrentQuery.State = 2 Then SQL.Add("AND HANDLE <> :HANDLE")
    Case 2
      SQL.Clear
      SQL.Add("SELECT HANDLE FROM SAM_CONTRATO WHERE CONTRATO = :CONTRATO AND EMPRESA = :EMPRESA")
      SQL.ParamByName("EMPRESA").Value = CurrentQuery.FieldByName("EMPRESA").AsInteger
      If CurrentQuery.State = 2 Then SQL.Add("AND HANDLE <> :HANDLE")
  End Select
  SQL.ParamByName("CONTRATO").Value = CurrentQuery.FieldByName("CONTRATO").AsInteger
  If CurrentQuery.State = 2 Then SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger

  SQL.Active = True
  If SQL.EOF Then
    NumeroContratoUnico = True
  Else
    NumeroContratoUnico = False
  End If
End Function


Public Function CheckNumeroAutomatico As Boolean
  Dim SQL2 As Object
  Set SQL2 = NewQuery
  Dim SQL3 As Object
  Set SQL3 = NewQuery

  SQL2.Add("SELECT COUNT(HANDLE) QTD FROM SAM_FAMILIA ")
  SQL2.Add(" WHERE CONTRATO = :CONTRATO               ")
  SQL2.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL2.Active = True


  SQL3.Add("SELECT COUNT(HANDLE) QTD FROM SAM_BENEFICIARIO ")
  SQL3.Add(" WHERE CONTRATO = :CONTRATO                    ")
  SQL3.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL3.Active = True

  'verifica a troca do tipo de códigos
  If NUMEROFAMILIAAUTOMATICO.Visible Then
    If(CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString = "S")And(vNumeroFamiliaAutomaticoAnterior = "N")Then
    If SQL2.FieldByName("QTD").AsInteger >0 Then
      bsShowMessage("Existem famílias cadastrados com número informado." + Chr(13) + "Impossível alterar o número da família para automático.", "E")
      CheckNumeroAutomatico = False
      Set SQL2 = Nothing
      Set SQL3 = Nothing
      Exit Function
    End If
  End If
End If
If NUMEROBENEFAUTOMATICO.Visible Then
  If(CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString = "S")And(vNumeroBenefAutomaticoAnterior = "N")Then
  If SQL3.FieldByName("QTD").AsInteger >0 Then
    bsShowMessage("Existem beneficiários cadastrados com número informado." + Chr(13) + "Impossível alterar o número do beneficiário para automático.", "E")
    CheckNumeroAutomatico = False
    Set SQL2 = Nothing
    Set SQL3 = Nothing
    Exit Function
  End If
End If
End If

Set SQL2 = Nothing
Set SQL3 = Nothing

CheckNumeroAutomatico = True

End Function

Public Function CriaContadorContrato

  Dim Chave As Long
  Dim Sequencia As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_PARAMETROSBENEFICIARIO")
  SQL.Active = True
  Select Case SQL.FieldByName("NUMEROCONTRATOUNICO").AsInteger
    Case 1
      Chave = 0 'ZERO irá funcionar como um contador universal para o sistema
    Case 2
      Chave = CurrentQuery.FieldByName("EMPRESA").AsInteger 'Cada EMPRESA é uma chave para o contador que não é universal
  End Select
  NewCounter("SAM_CONTRATO", Chave, 1, Sequencia)
  CriaContadorContrato = Sequencia
End Function

Public Function HerdaParametro
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT NUMEROFAMILIAAUTOMATICO, NUMEROBENEFAUTOMATICO FROM EMPRESAS WHERE HANDLE =:EMPRESA")
  SQL.ParamByName("EMPRESA").AsInteger = CurrentQuery.FieldByName("EMPRESA").AsInteger
  SQL.Active = True

  CurrentQuery.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString = SQL.FieldByName("NUMEROFAMILIAAUTOMATICO").AsString
  CurrentQuery.FieldByName("NUMEROBENEFAUTOMATICO").AsString = SQL.FieldByName("NUMEROBENEFAUTOMATICO").AsString
  CurrentQuery.UpdateRecord

End Function


Public Sub BOTAOCONTRATO_OnClick()

  Dim Interface As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("A tabela não pode estar em edição", "E")
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SAMConsultaBenef.Consultas")
  Interface.Executar(CurrentSystem, 1, CurrentQuery.FieldByName("HANDLE").AsInteger, 0, 0)
  Set Interface = Nothing

End Sub

Public Function AtualizaRotulosEndereco( pEnderecoCorrespondencia As Long)
	If (CurrentQuery.FieldByName("TABTIPOCONTRATO").AsInteger = 1) Then
		Dim vQryEndereco As Object
		Dim vListaEnderecos As String
		vListaEnderecos = ""

		If pEnderecoCorrespondencia > 0 Then
			On Error GoTo Except
				Dim vLogradouro, vNumero, vComplemento, vBairro, vCEP, vTelefone1, vTelefone2, vFax, vCelular, vRamal As String
				Dim vMunicipio, vEstado, vTipoLogradouro As String

				Set vQryEndereco = NewQuery
				vQryEndereco.Active = False
				vQryEndereco.Clear
				vQryEndereco.Add("SELECT E.HANDLE, E.ESTADO, E.MUNICIPIO, E.BAIRRO, E.CEP, E.NUMERO, E.COMPLEMENTO, E.TELEFONE1, ")
				vQryEndereco.Add("       E.TELEFONE2, E.FAX, E.LOGRADOURO, E.CELULAR, E.RAMAL, LT.DESCRICAO AS TIPOLOGRADOURO,   ")
				vQryEndereco.Add("       ES.NOME NOMEESTADO, M.NOME NOMEMUNICIPIO ")
				vQryEndereco.Add("  FROM SAM_ENDERECO E")
				vQryEndereco.Add("  LEFT JOIN LOGRADOUROS_TIPO LT ON LT.HANDLE = E.TIPOLOGRADOURO ")
				vQryEndereco.Add("  LEFT JOIN ESTADOS ES ON ES.HANDLE = E.ESTADO ")
				vQryEndereco.Add("  LEFT JOIN MUNICIPIOS M ON M.HANDLE = E.MUNICIPIO ")
				vQryEndereco.Add(" WHERE E.HANDLE = :HANDLE ")
				vQryEndereco.ParamByName("HANDLE").AsInteger = pEnderecoCorrespondencia
				vQryEndereco.Active = True

				If (vQryEndereco.FieldByName("HANDLE").AsInteger > 0) Then
					vLogradouro =	preencheValor(""			, vQryEndereco.FieldByName("LOGRADOURO").AsString, 		""		,"")
					vNumero     =	preencheValor(", Nº "		, vQryEndereco.FieldByName("NUMERO").AsString, 			""		,"")
					vComplemento=	preencheValor("Complemento: ",vQryEndereco.FieldByName("COMPLEMENTO").AsString, 	"     "	,"")
					vBairro 	=	preencheValor("Bairro: "	, vQryEndereco.FieldByName("BAIRRO").AsString, 			""		,"")
					vCEP 		=	preencheValor("CEP: "		, vQryEndereco.FieldByName("CEP").AsString, 			"     "	,"")
					vMunicipio	=	preencheValor("Município: "	, vQryEndereco.FieldByName("NOMEMUNICIPIO").AsString,	"     "	,"")
					vEstado		=	preencheValor("Estado: "	, vQryEndereco.FieldByName("NOMEESTADO").AsString, 		""		,"")
					vTelefone1	=	preencheValor("Telefone 1: ", vQryEndereco.FieldByName("TELEFONE1").AsString, 		"     "	,"")
					vTelefone2	=	preencheValor("Telefone 2: ", vQryEndereco.FieldByName("TELEFONE2").AsString, 		"     "	,"")
					vRamal		=	preencheValor("Ramal: "		, vQryEndereco.FieldByName("RAMAL").AsString, 			"     "	,"")
					vFax		=	preencheValor("Fax: "		, vQryEndereco.FieldByName("FAX").AsString, 			""		,"")
					vCelular	=	preencheValor("Celular: "	, vQryEndereco.FieldByName("CELULAR").AsString, 		"     "	,"")
					vTipoLogradouro = preencheValor(""			, vQryEndereco.FieldByName("TIPOLOGRADOURO").AsString,	" ",	  	 "")

					preencheRotulosEndereco("COR", vTipoLogradouro + vLogradouro, _
												   vNumero, _
												   vComplemento + vBairro, _
												   vCEP + vMunicipio + vEstado, _
												   vTelefone1 + vTelefone2 + vRamal + vFax, _
												   IIf( vCelular <> "", vCelular + "     ", ""))
				Else
					preencheRotulosEndereco("COR", "", "", "", "", "", "")
				End If
				vQryEndereco.Active = False
				Set vQryEndereco = Nothing
				Exit Function
			Except:
				Set vQryEndereco = Nothing
				Err.Raise(Err.Number, Err.Source, "Falha ao exibir endereços do Contrato: " + Err.Description)
		Else
			preencheRotulosEndereco("COR", "", "", "", "", "", "")
		End If
	End If
End Function

Public Function preencheValor(pPrefixo As String, pValor As String, pSufixo As String, pValorSeVazio) As String
	If pValor <> "" Then
		preencheValor = pPrefixo + pValor + pSufixo
	Else
		preencheValor = pValorSeVazio
	End If
End Function

Public Function preencheRotulosEndereco(pTipo As String, pRot0 As String, pRot1 As String, pRot2 As String, pRot3 As String, pRot4 As String, pRot5 As String)
	Select Case pTipo
		Case "COR"
		  	ROTULOCORRESP1.Text = pRot0 + pRot1
		  	ROTULOCORRESP2.Text = pRot2
			ROTULOCORRESP3.Text = pRot3
		 	ROTULOCORRESP4.Text = pRot4
		 	ROTULOCORRESP5.Text = pRot5
	End Select
End Function

Public Function GravaEndereco( pContrato As Long, pEnderecoCorrespondencia As Long)
	Dim iniciouTransacao As Boolean
	Dim msgErro As String
	Dim numErro As Long

	On Error GoTo Except
		Dim sqlUp As Object
		Set sqlUp = NewQuery

		If Not pContrato > 0 Then
			Err.Raise(1, Err, "Falha ao atualizar Endereço: Contrato não informado!")
		End If
		Dim vEndCor As String
		If pEnderecoCorrespondencia > 0 Then
			vEndCor = CStr(pEnderecoCorrespondencia)
		Else
			vEndCor = "NULL"
		End If

		sqlUp.Add("UPDATE SAM_CONTRATO SET ENDERECOCORRESPONDENCIA = " + vEndCor + " WHERE HANDLE = :HANDLE ")
		sqlUp.ParamByName("HANDLE").AsInteger = pContrato

		iniciouTransacao = False
		If Not InTransaction Then
			StartTransaction
			iniciouTransacao = True
		End If

		sqlUp.ExecSQL

		If iniciouTransacao And InTransaction Then
			Commit
			iniciouTransacao = False
		End If

		Set sqlUp = Nothing
		Exit Function
	Except:
		msgErro = Err.Description
		numErro = Err.Number
		Set sqlUp = Nothing
		If iniciouTransacao And InTransaction Then
			Rollback
		End If
		Err.Raise(numErro, Err, msgErro)
End Function
