'HASH: 250C4ADA652E42D3DF024E8C6F011CA8
'Macro: SAM_MIGRACAO_PROCESSO
'#Uses "*bsShowMessage"

Dim vPodeIncluirInterfaceMigracao As Boolean 'sms 31310

Public Function VerificaDataFechamento()As Boolean

  Dim qFechamento
  Set qFechamento = NewQuery
  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  qFechamento.Active = True


  Dim vMesComp As Integer
  Dim vAnoComp As Integer

  Dim vMesFechamento As Integer
  Dim vAnoFechamento As Integer



  vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
  vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)

  vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
  vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

  If(vAnoComp <vAnoFechamento)Or _
     (vAnoComp = vAnoFechamento And vMesComp <vMesFechamento)Then
  bsShowMessage("A competência não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  VerificaDataFechamento = False
  Exit Function
End If

If CurrentQuery.FieldByName("DATAMIGRACAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
  VerificaDataFechamento = False
  bsShowMessage("Não é possível migrar beneficiários com data de migração inferior a data de fechamento - Parâmetros Gerais", "E")
  Exit Function
End If

If CurrentQuery.FieldByName("DATACONTABIL").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
  VerificaDataFechamento = False
  bsShowMessage("Não é possível migrar contratos com data contábil inferior a data de fechamento - Parâmetros Gerais", "E")
  Exit Function
End If

Set qFechamento = Nothing
VerificaDataFechamento = True

End Function


Public Sub BENEFICIARIO_OnPopup(ShowPopup As Boolean)
  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTexto As String
  Dim vTipopesquisa As Integer
  vTipopesquisa=2

  vTexto = BENEFICIARIO.LocateText

  If IsNumeric(vTexto) Then
    vTipopesquisa=1
  End If

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "BENEFICIARIO|NOME|CONTRATO|CARTAO"

  vCriterio = " VIACARTAO IS NOT NULL AND DVCARTAO IS NOT NULL  "
  vCriterio = vCriterio + " AND DATACANCELAMENTO IS NULL        "

  If Not CurrentQuery.FieldByName("CONTRATOORIGEM").IsNull Then
    vCriterio = vCriterio + "AND CONTRATO = " + Str(CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger)
  End If

  vCampos = "Beneficiário|Nome|Contrato|Cartão"

  vHandle = Interface.Exec(CurrentSystem, "SAM_BENEFICIARIO", vColunas, vTipopesquisa, vCampos, vCriterio, "Contratos", True, vTexto)

  If vHandle <>0 Then
    CurrentQuery.FieldByName("BENEFICIARIO").Value = vHandle
    Set SQL = NewQuery
    SQL.Clear
    SQL.Add("SELECT CONTRATO               ")
    SQL.Add("FROM SAM_BENEFICIARIO         ")
    SQL.Add("WHERE HANDLE = :HBENEFICIARIO ")
    SQL.ParamByName("HBENEFICIARIO").Value = vHandle
    SQL.Active = True

    If Not SQL.EOF Then
      SQL.First
      CurrentQuery.FieldByName("CONTRATOORIGEM").Value=SQL.FieldByName("CONTRATO").AsInteger
    End If

  End If

  Set Interface = Nothing

End Sub

Public Sub BOTAOGERAR_OnClick()

  If VerificaDataFechamento = True Then
    If(CurrentQuery.State = 1)And(CurrentQuery.FieldByName("DATAGERACAO").IsNull)Then

 		If VisibleMode Then

    	    Dim Interface As Object
	        Set Interface = CreateBennerObject("CONTRATO.Migracao")
        	bsShowMessage(Interface.Gerar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger), "I")
        	Set Interface = Nothing
        Else
	         Dim vsMensagemErro As String
   			 Dim viRetorno As Long


	         Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    		 viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
        			                             "CONTRATO", _
                    			                 "Migracao_Gerar", _
                                			     "Processo de Migração de Beneficiários - Geração : " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                    			 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     			 "SAM_MIGRACAO_PROCESSO", _
                                     			 "SITUACAOGERAR", _
                                     			 "", _
                                    			 "", _
                                     			 "P", _
                                     			 False, _
                                      			 vsMensagemErro, _
                                     			 Null)


        	If viRetorno = 0 Then
  				bsShowMessage("Processo enviado para processamento no servidor, ver ocorrências "+ Chr(13) + vsMensagemErro, "I")
	 		Else
    			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
       		End If

       End If

       WriteAudit("G", HandleOfTable("SAM_MIGRACAO_PROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Migração de Beneficiários - Geração")

    Else
  		bsShowMessage("Rotina já gerada ou processada", "I")  ' SMS 92204 - Paulo Melo - 11/01/2008
    End If
End If

End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  If VerificaDataFechamento = True Then
    If(CurrentQuery.State = 1)And(Not CurrentQuery.FieldByName("DATAGERACAO").IsNull) And (CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull) Then


    	If VisibleMode Then

   	    	Dim Interface As Object
   			Set Interface = CreateBennerObject("CONTRATO.Migracao")
   			bsShowMessage(Interface.Processar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, True), "I")
   			Set Interface = Nothing

        Else
	    	Dim vsMensagemErro As String
   			Dim viRetorno As Long

	        Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    		viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
        		                             "CONTRATO", _
                   			                 "Migracao_Processar", _
                               			     "Processo de Migração de Beneficiários - Processamento : " + CStr(CurrentQuery.FieldByName("HANDLE").AsInteger), _
                                   			 CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                   			 "SAM_MIGRACAO_PROCESSO", _
                                   			 "SITUACAOPROCESSAR", _
                                   			 "", _
                                   			 "", _
                                   			 "P", _
                                   			 False, _
                                   			 vsMensagemErro, _
                                   			 Null)


        	If viRetorno = 0 Then
  				bsShowMessage("Processo enviado para processamento no servidor, ver ocorrências "+ Chr(13) + vsMensagemErro, "I")
	 		Else
    			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
       		End If

       End If

       WriteAudit("P", HandleOfTable("SAM_MIGRACAO_PROCESSO"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Processo de Migração de Beneficiários - Processamento")

    Else
    		bsShowMessage("Rotina ainda não gerada ou já processada", "I")  ' SMS 91312 - Paulo Melo - 11/01/2008
  	End If
End If

End Sub

Public Sub CONTRATODESTINO_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim vTexto As String
  Dim vTipopesquisa As Integer
  vTipopesquisa=2

  vTexto = CONTRATOORIGEM.LocateText

  If IsNumeric(vTexto) Then
    vTipopesquisa=1
  End If

  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"

  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND NAOINCLUIRBENEFICIARIO = 'N' "
  vCriterio = vCriterio + "AND HANDLE <> " + Str(CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger)
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vTipopesquisa, vCampos, vCriterio, "Contratos", True, vTexto)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATODESTINO").Value = vHandle
  End If

  Set Interface = Nothing

End Sub

Public Sub CONTRATOORIGEM_OnPopup(ShowPopup As Boolean)

  Dim Interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String
  Dim pesquisarPor As Integer
  Dim vTexto As String
  Dim vTipopesquisa As Integer
  vTipopesquisa=2

  vTexto = CONTRATOORIGEM.LocateText

  If IsNumeric(vTexto) Then
    vTipopesquisa=1
  End If


  ShowPopup = False
  Set Interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"
  vCriterio = "DATACANCELAMENTO IS NULL "
  vCriterio = vCriterio + "AND HANDLE <> " + Str(CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger)
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = Interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, vTipopesquisa, vCampos, vCriterio, "Contratos", True, vTexto)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOORIGEM").Value = vHandle

    If Not CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
       Set SQL = NewQuery
       SQL.Clear
       SQL.Add("SELECT *                        ")
       SQL.Add("FROM SAM_BENEFICIARIO           ")
       SQL.Add("WHERE HANDLE   = :HBENEFICIARIO ")
       SQL.Add("  AND CONTRATO = :HCONTRATO     ")
       SQL.ParamByName("HBENEFICIARIO").Value = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
       SQL.ParamByName("HCONTRATO").Value = vHandle
       SQL.Active = True

       If SQL.EOF Then
         CurrentQuery.FieldByName("BENEFICIARIO").Clear
       End If
    End If

  End If

  Set Interface = Nothing

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim SQL As Object
  Dim SQL2 As Object
  Dim SQLDel As Object

  If Not CurrentQuery.FieldByName("DATAPROCESSAMENTO").IsNull Then
    bsShowMessage("Operação cancelada. A rotina já foi processada", "I")
  Else

    If CurrentUser = CurrentQuery.FieldByName("RESPONSAVEL").AsInteger Then

      Set SQL = NewQuery
      Set SQL2 = NewQuery
      Set SQLDel = NewQuery

      SQL.Clear
      SQL.Add("SELECT HANDLE")
      SQL.Add("FROM SAM_MIGRACAO_PROCESSOBENEF")
      SQL.Add("WHERE MIGRACAOPROCESSO = :HMIGRACAOPROCESSO")
      SQL.ParamByName("HMIGRACAOPROCESSO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQL.Active = True

      While Not SQL.EOF
        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_CAR")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_LIM")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_PF")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_MOD")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_FRQ")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_BENEF_LIMITACAO")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_BENEF_FRANQUIA")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQL2.Clear
        SQL2.Add("SELECT HANDLE")
        SQL2.Add("FROM SAM_MIGRACAO_BENEF_PFEVENTO")
        SQL2.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQL2.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQL2.Active = True

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_BENEF_PFEVENTO_FX")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEFPFEVENTO = :HMIGRACAOPROCESSOBENEFPFEVENTO")

        While Not SQL2.EOF
          SQLDel.ParamByName("HMIGRACAOPROCESSOBENEFPFEVENTO").Value = SQL2.FieldByName("HANDLE").AsInteger
          SQLDel.ExecSQL

          SQL2.Next
        Wend

        SQLDel.Clear
        SQLDel.Add("DELETE FROM SAM_MIGRACAO_BENEF_PFEVENTO")
        SQLDel.Add("WHERE MIGRACAOPROCESSOBENEF = :HMIGRACAOPROCESSOBENEF")
        SQLDel.ParamByName("HMIGRACAOPROCESSOBENEF").Value = SQL.FieldByName("HANDLE").AsInteger
        SQLDel.ExecSQL

        SQL.Next
      Wend

      Set SQL = Nothing
      Set SQL2 = Nothing

      SQLDel.Clear
      SQLDel.Add("DELETE FROM SAM_MIGRACAO_PROCESSOBENEF")
      SQLDel.Add("WHERE MIGRACAOPROCESSO = :HMIGRACAOPROCESSO")
      SQLDel.ParamByName("HMIGRACAOPROCESSO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
      SQLDel.ExecSQL

      Set SQLDel = Nothing

    Else

      CanContinue = False
      bsShowMessage("Operação cancelada. Usuário não é o Responsável", "E")

    End If
  End If

End Sub


Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vPodeIncluirInterfaceMigracao = True 'sms 31310
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  vPodeIncluirInterfaceMigracao = False 'sms 31310
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qFechamento
  Set qFechamento = NewQuery
  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  qFechamento.Active = True



  'sms 60045
  If CurrentQuery.FieldByName("TABTIPOMIGRACAO").AsInteger = 4 Then
    If (CurrentQuery.FieldByName("DIASCANCFUTURO").AsInteger > 0) And (CurrentQuery.FieldByName("MOTIVOCANCFUTURO").IsNull) Then
      bsShowMessage("Motivo cancelamento futuro deve ser informado", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DIASCANCFUTURO").AsInteger = 0) And (Not CurrentQuery.FieldByName("MOTIVOCANCFUTURO").IsNull) And (CurrentQuery.FieldByName("PRAZOCANCFUTURO").AsInteger = 1) Then
      bsShowMessage("Dias para cancelamento futuro deve ser informado", "E")
      CanContinue = False
      Exit Sub
    End If

    If CurrentQuery.FieldByName("DATAINICIALCANCFUTURO").AsDateTime < ServerDate Then
      bsShowMessage("Informar data cancelamento inicial maior ou igual a data atual", "E")
      CanContinue = False
      Exit Sub
    End If

    If CurrentQuery.FieldByName("DATAFINALCANCFUTURO").AsDateTime < CurrentQuery.FieldByName("DATAINICIALCANCFUTURO").AsDateTime Then
      bsShowMessage("Data cancelamento final deve ser maior ou igual a data cancelamento inicial", "E")
      CanContinue = False
      Exit Sub
    End If


  End If

  If CurrentQuery.State = 3 Then


    Dim vMesComp As Integer
    Dim vAnoComp As Integer

    Dim vMesFechamento As Integer
    Dim vAnoFechamento As Integer

    vMesComp = DatePart("m", CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)
    vAnoComp = DatePart("yyyy", CurrentQuery.FieldByName("COMPETENCIA").AsDateTime)

    vMesFechamento = DatePart("m", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)
    vAnoFechamento = DatePart("yyyy", qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime)

    If(vAnoComp <vAnoFechamento)Or _
       (vAnoComp = vAnoFechamento And vMesComp <vMesFechamento)Then
    CanContinue = False
    bsShowMessage("A competência não pode ser inferior à data de fechamento - Parâmetros Gerais", "E")
  End If

  If CurrentQuery.FieldByName("DATAMIGRACAO").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    bsShowMessage("Não é possível migrar beneficiários com data de migração inferior a data de fechamento - Parâmetros Gerais", "E")
    CanContinue = False
  End If

  If CurrentQuery.FieldByName("DATACONTABIL").AsDateTime <qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
    bsShowMessage("Não é possível migrar contratos com data contábil inferior a data de fechamento - Parâmetros Gerais", "E")
    CanContinue = False
  End If
End If

Set qFechamento = Nothing

If CurrentQuery.FieldByName("CONTRATOORIGEM").AsInteger = CurrentQuery.FieldByName("CONTRATODESTINO").AsInteger Then
  CanContinue = False
  bsShowMessage("Contrato origem não pode ser igual ao destino!", "E")
  Exit Sub
End If

'sms 31310
If (Not vPodeIncluirInterfaceMigracao And CurrentQuery.FieldByName("TABTIPOMIGRACAO").AsInteger = 3) Then
  CanContinue = False
  bsShowMessage("Opção 'Interface de migração' pode ser utilizada somente pela Interface de Migração!", "E")
  Exit Sub
End If
'fim sms 31310

If CurrentQuery.FieldByName("TABTIPOMIGRACAO").AsInteger = 5 Then

  If (Not (CurrentQuery.FieldByName("MOTIVOCANCFUTURO").IsNull)) Or (Not(CurrentQuery.FieldByName("DATACANCFUTUROBENEF").IsNull)) Then
    If (CurrentQuery.FieldByName("MOTIVOCANCFUTURO").IsNull) Then
      bsShowMessage("Motivo cancelamento futuro deve ser informado", "E")
      CanContinue = False
      Exit Sub
    End If

    If (CurrentQuery.FieldByName("DATACANCFUTUROBENEF").IsNull) Then
      bsShowMessage("Data para cancelamento futuro deve ser informado", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

    If CurrentQuery.FieldByName("BENEFICIARIO").IsNull Then
      bsShowMessage("O beneficiário deve ser informado!", "E")
      CanContinue = False
      Exit Sub
    End If
End If

End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("COMPETENCIA").Value = DateAdd("m", 1, ServerDate)
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOGERAR"
			BOTAOGERAR_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
	End Select
End Sub
