'HASH: 674694E4D8521A69044CC5C98A2E73D6
'Macro: SAM_FAMILIA_MOD
'#Uses "*UltimoDiaCompetencia"
'#Uses "*bsShowMessage"
Dim Voperadora As Integer
Dim SqlOperadora
Dim Sql
Dim vCompFinal As Date

Public Sub BOTAOCANCELAR_OnClick()

  If CurrentQuery.State = 1 Then
    Dim Interface As Object

  If VisibleMode Then

    If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then

    		Dim vsMensagemErro As String
			Dim viRetorno As Integer
			Dim vvContainer As CSDContainer

	    	Set vvContainer = NewContainer

	    	Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")

   			viRetorno = Interface.Exec(CurrentSystem, _
  										1, _
	                                   	"TV_FORM0018", _
       	                        		"Cancelamento do Módulo da Família", _
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



    Else
    	bsShowMessage("Cancelamento já processado!","I")
    End If

  End If
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOREATIVAR_OnClick()
	If VisibleMode Then
  		If CurrentQuery.State = 1 Then
    		If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then


    			Dim vsMensagemErro As String
	    		Dim viRetorno As Integer
	    		Dim vvContainer As CSDContainer

			    Set vvContainer = NewContainer

			    Set Interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")


		   		viRetorno = Interface.Exec(CurrentSystem, _
	   									1, _
                                    	"TV_FORM0012", _
           	                        	"Reativação do Módulo da Família", _
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
    		Else
    			bsShowMessage("Cancelamento não processado!", "I")
    		End If
    		CurrentQuery.Active = False
    		CurrentQuery.Active = True
  		End If
  	End If
End Sub

Public Sub MODULO_OnChange()
  If CurrentQuery.State = 3 Then ' Inclusão
    Dim qm As Object
    Set qm = NewQuery
    qm.Clear
    qm.Add("SELECT CM.PRIMEIRAPARCELA, CM.PARCELADIAS, CM.AGENTEAGENCIAVENDAS, CM.TIPOCOMISSAO")
    qm.Add("FROM SAM_CONTRATO_MOD CM")
    qm.Add("WHERE CM.HANDLE = :CONTRATOMOD")
    qm.ParamByName("CONTRATOMOD").Value = CurrentQuery.FieldByName("MODULO").AsInteger
    qm.Active = True
    If Not qm.EOF Then
      If(Not qm.FieldByName("PRIMEIRAPARCELA").IsNull)Then
      CurrentQuery.FieldByName("PRIMEIRAPARCELA").Value = qm.FieldByName("PRIMEIRAPARCELA").Value
    End If
    If(Not qm.FieldByName("PARCELADIAS").IsNull)Then
    CurrentQuery.FieldByName("PARCELADIAS").Value = qm.FieldByName("PARCELADIAS").Value
  Else
    CurrentQuery.FieldByName("PARCELADIAS").Clear
  End If
  If(Not qm.FieldByName("AGENTEAGENCIAVENDAS").IsNull)Then
  CurrentQuery.FieldByName("AGENTEAGENCIAVENDAS").Value = qm.FieldByName("AGENTEAGENCIAVENDAS").Value
Else
  CurrentQuery.FieldByName("AGENTEAGENCIAVENDAS").Clear
End If
If(Not qm.FieldByName("TIPOCOMISSAO").IsNull)Then
CurrentQuery.FieldByName("TIPOCOMISSAO").Value = qm.FieldByName("TIPOCOMISSAO").Value
Else
  CurrentQuery.FieldByName("TIPOCOMISSAO").Clear
End If
End If
End If
qm.Active = False
Set qm = Nothing
End Sub



Public Sub TABLE_AfterScroll()

  If CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    AGENTEAGENCIAVENDAS.ReadOnly = False
    TIPOCOMISSAO.ReadOnly = False
    PARCELADIAS.ReadOnly = False
    PRIMEIRAPARCELA.ReadOnly = False
  Else
    AGENTEAGENCIAVENDAS.ReadOnly = True
    TIPOCOMISSAO.ReadOnly = True
    PARCELADIAS.ReadOnly = True
    PRIMEIRAPARCELA.ReadOnly = True
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    RecordHandleOfTable("SAM_FAMILIA"), _
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
  	MODULO.WebLocalWhere = "A1.HANDLE IN (SELECT SAM_CONTRATO_MOD.MODULO FROM SAM_CONTRATO_MOD WHERE SAM_CONTRATO_MOD.CONTRATO = CONTRATO)"
  ElseIf VisibleMode Then
  	MODULO.LocalWhere = "SAM_MODULO.HANDLE IN (SELECT SAM_CONTRATO_MOD.MODULO FROM SAM_CONTRATO_MOD WHERE SAM_CONTRATO_MOD.CONTRATO = CONTRATO)"
  End If

  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    RecordHandleOfTable("SAM_FAMILIA"), _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido editar o módulo por motivo de suspensão!", "E")
    CanContinue = False
    CurrentQuery.Cancel
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  If WebMode Then
  	MODULO.WebLocalWhere = "A1.HANDLE IN (SELECT SAM_CONTRATO_MOD.MODULO FROM SAM_CONTRATO_MOD WHERE SAM_CONTRATO_MOD.CONTRATO = CONTRATO)"
  ElseIf VisibleMode Then
  	MODULO.LocalWhere = "SAM_MODULO.HANDLE IN (SELECT SAM_CONTRATO_MOD.MODULO FROM SAM_CONTRATO_MOD WHERE SAM_CONTRATO_MOD.CONTRATO = CONTRATO)"
  End If

  'Verifica suspensão -Juliano 09-12-02----------------------------------------------------------------------------------------------
  Dim vDataFinalSuspensao As Date
  Dim BSBen001Dll As Object
  Set BSBen001Dll = CreateBennerObject("BSBen001.Beneficiario")
  If BSBen001Dll.VerificaSuspensao(CurrentSystem, _
                                    0, _
                                    RecordHandleOfTable("SAM_FAMILIA"), _
                                    RecordHandleOfTable("SAM_CONTRATO"), _
                                    vDataFinalSuspensao)Then
    bsShowMessage("Não é permitido inserir o módulo por motivo de suspensão!", "E")
    CanContinue = False
    CurrentQuery.Cancel
    Exit Sub
  End If
  Set BSBen001Dll = Nothing
  '------------------------------------------------------------------------------------------------------------------------------------
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.FieldByName("SEGUNDAPARCELA").AsString = "2" And _
                              CurrentQuery.FieldByName("PRIMEIRAPARCELA").AsString <>"2" Then
    CanContinue = False
    bsShowMessage("Para segunda parcela 'Proporcional' a primeira parcela deve ser integral", "E")
    Exit Sub
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
        SqlOperadora.Add("SELECT MAX(DATAFINAL)DATAFINAL    ")
        SqlOperadora.Add("  FROM SAM_MSPROCESSO             ")
        SqlOperadora.Add(" WHERE OPERADORA = :OPERADORA     ")
        SqlOperadora.Add("   AND SITUACAO  = 'P'            ")
        SqlOperadora.Add("   AND TABTIPOOPERADORA = :TABTIPO")
        SqlOperadora.Add("   AND DATAEXPORTACAO IS NOT NULL ")
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
  'FiM ANDERSON
  '--------------------------------------------------------------------------------------------------------------------------

  Dim qm As Object

  Set qm = NewQuery

  qm.Clear
  qm.Add("SELECT DATAADESAO")
  qm.Add("FROM SAM_FAMILIA")
  qm.Add("WHERE HANDLE = :HFAMILIA")
  qm.ParamByName("HFAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
  qm.Active = True

  If CurrentQuery.FieldByName("DATAADESAO").AsDateTime <qm.FieldByName("DATAADESAO").AsDateTime Then
    CanContinue = False
    bsShowMessage("Data de adesão inferior à adesão da família", "E")
    Set qm = Nothing
    Exit Sub
  End If

  qm.Clear
  qm.Add("SELECT DATAADESAO")
  qm.Add("FROM SAM_CONTRATO_MOD")
  qm.Add("WHERE HANDLE = :HCONTRATOMOD")
  qm.ParamByName("HCONTRATOMOD").Value = CurrentQuery.FieldByName("MODULO").AsInteger
  qm.Active = True

  If CurrentQuery.FieldByName("DATAADESAO").AsDateTime <qm.FieldByName("DATAADESAO").AsDateTime Then
    CanContinue = False
    bsShowMessage("Data de adesão inferior à adesão do módulo no contrato", "E")
    Set qm = Nothing
    Exit Sub
  End If

  Set qm = Nothing


  Dim Interface As Object
  Dim Linha As String
  Dim Condicao As String

  Condicao = " AND MODULO = " + CurrentQuery.FieldByName("MODULO").AsString

  Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

  Linha = Interface.Vigencia(CurrentSystem, "SAM_FAMILIA_MOD", "DATAADESAO", "DATACANCELAMENTO", CurrentQuery.FieldByName("DATAADESAO").AsDateTime, CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime, "FAMILIA", Condicao)

  If Linha = "" Then
    CanContinue = True
  Else
    CanContinue = False
    bsShowMessage(Linha, "E")
    Set Interface = Nothing
    Exit Sub
  End If

  If(CurrentQuery.FieldByName("PRIMEIRAPARCELA").AsString <>"3")And _
     (CurrentQuery.FieldByName("SEGUNDAPARCELA").AsString = "1")Then ' isento ou integral
  If Not CurrentQuery.FieldByName("PARCELADIAS").IsNull Then
    bsShowMessage("Informar quantidade de dias somente para primeira parcela proporcional!", "E")
    CanContinue = False
    Exit Sub
  End If
End If
CanContinue = CheckVigencia


Dim qFechamento As Object
Set qFechamento = NewQuery

qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
qFechamento.Active = True

If CurrentQuery.State = 3 Then
  If CurrentQuery.FieldByName("DATAADESAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    CanContinue = False
    bsShowMessage("Não é possível cadastrar data de adesão inferior a data de fechamento - Parâmetros Gerais", "E")
  End If
End If

If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
  If CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    CanContinue = False
    bsShowMessage("Não é possível cadastrar data de cancelamento inferior a data de fechamento - Parâmetros Gerais", "E")
  End If
End If

Set qFechamento = Nothing

End Sub


Public Function CheckVigencia As Boolean
  CheckVigencia = True
  Dim Sql As Object
  Set Sql = NewQuery
  Sql.Add("SELECT * FROM SAM_FAMILIA WHERE HANDLE = :FAMILIA")
  Sql.ParamByName("FAMILIA").Value = CurrentQuery.FieldByName("FAMILIA").AsInteger
  Sql.Active = True
  If CurrentQuery.FieldByName("DATAADESAO").AsDateTime <Sql.FieldByName("DATAADESAO").AsDateTime Then
    bsShowMessage("Data de adesão do Módulo menor que a Adesão da Família!", "E")
    CheckVigencia = False
  Else
    If Not Sql.FieldByName("DATACANCELAMENTO").IsNull Then
      If CurrentQuery.FieldByName("DATAADESAO").AsDateTime >Sql.FieldByName("DATACANCELAMENTO").AsDateTime Then
        bsShowMessage("Data de adesão do módulo maior que o cancelamento da Família!", "E")
        CheckVigencia = False
      End If
    End If
  End If
  If Not CurrentQuery.FieldByName("DATACANCELAMENTO").IsNull Then
    If CurrentQuery.FieldByName("DATAADESAO").AsDateTime >CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime Then
      bsShowMessage("Data de Adesão maior que data de cancelamento do módulo!", "E")
      CheckVigencia = False
    End If
  End If
  Set Sql = Nothing
End Function




Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOREATIVAR"
			BOTAOREATIVAR_OnClick
	End Select
End Sub
