'HASH: 5028F0C4E80D4B1A13A09DEB3B7908F0
'#USES "*CriaTabelaTemporariaSqlServer"

  Dim  gBeneficiario     As  Long
  Dim  gLocal            As  Long
  Dim  gRegime           As  Long
  Dim  gCondicao         As  Long
  Dim  gTipo             As  Long
  Dim  gObjetivo         As  Long
  Dim  gFinalidade       As  Long
  Dim  gConvenio         As  Long
  Dim  gCodpagto         As  Long
  Dim  gPlano            As  Long
  Dim  gXTHM             As  Long


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Interface          As Object
  Dim  SQL               As Object
  Dim  PS                As Object
  Dim  pBeneficiario     As  Long
  Dim  pLocal            As  Long
  Dim  pRegime           As  Long
  Dim  pCondicao         As  Long
  Dim  pTipo             As  Long
  Dim  pObjetivo         As  Long
  Dim  pFinalidade       As  Long
  Dim  pRecebedor        As  Long
  Dim  pLocalExecucao    As  Long
  Dim  pMunicipio        As  Long
  Dim  pEstado           As  Long
  Dim  pConvenio         As  Long
  Dim  pCodpagto         As  Long
  Dim  pXTHM             As  Long
  Dim  pAcomodacao       As  Long
  Dim  pEvento           As  Long
  Dim  pGrau             As  Long
  Dim  pTecnicaCirurgica As  String
  Dim  pData             As  Date
  Dim  pValorEvento      As  Double
  Dim  pValorPF          As  Double
  Dim  pQuantidade       As  Long
  Dim  pOcorrencias      As  String
  Dim  vDescEvento       As  String
  Dim  pPLano            As  Long
  Dim  pGrauEHMatMed     As  String


 If Not InTransaction Then
   StartTransaction
 End If

 If InStr(SQLServer, "SQL") > 0 Then
 'If InStr(1,"SQL",SQLServer) > 0 Then
   Dim SQLx As Object
   Set SQLx = NewQuery

   On Error GoTo TabelasTemporarias

     SQLx.Clear
     SQLx.Add("SELECT 1 FROM #TMP_ORIGEMCALCULO")
     SQLx.Active = True

     Set SQLx = Nothing

     GoTo Procedure

   TabelasTemporarias:
     CriaTabelaTemporariaSqlServer
 End If


 Procedure:

 If InTransaction Then
   Commit
 End If

 pBeneficiario  = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
 pLocal         = CurrentQuery.FieldByName("LOCALATENDIMENTO").AsInteger
 pRegime        = CurrentQuery.FieldByName("REGIMEATENDIMENTO").AsInteger
 pCondicao      = CurrentQuery.FieldByName("CONDATENDIMENTO").AsInteger
 pTipo          = CurrentQuery.FieldByName("TIPOTRATAMENTO").AsInteger
 pObjetivo      = CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsInteger
 pFinalidade    = CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").AsInteger
 pRecebedor     = CurrentQuery.FieldByName("RECEBEDOR").AsInteger
 pLocalExecucao = CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger
 pMunicipio     = CurrentQuery.FieldByName("MUNICIPIO").AsInteger
 pEstado        = CurrentQuery.FieldByName("ESTADO").AsInteger
 pConvenio      = CurrentQuery.FieldByName("CONVENIO").AsInteger
 pCodpagto      = CurrentQuery.FieldByName("CODPAGAMENTO").AsInteger
 pXTHM          = CurrentQuery.FieldByName("XTHM").AsInteger
 pAcomodacao    = CurrentQuery.FieldByName("ACOMODACAO").AsInteger
 pEvento        = CurrentQuery.FieldByName("PROCEDIMENTO").AsInteger
 pGrau          = CurrentQuery.FieldByName("GRAU").AsInteger
 pData          = CurrentQuery.FieldByName("DATA").AsDateTime
 pQuantidade    = CurrentQuery.FieldByName("QUANTIDADE").AsInteger
 pPLano         = CurrentQuery.FieldByName("PLANO").AsInteger
 pTecnicaCirurgica = "1"
 '----------------------------------------------------------------------
 If	(pBeneficiario <= 0) And (pConvenio <= 0) Then
	If Not VisibleMode Then
		CancelDescription = "O beneficiário não foi informado. Informe o convênio. "
		CanContinue = False
	End If
 End If
 '----------------------------------------------------------------------
 ' pOcorrencias = "PLano = "+Str(pPLano) - "+Str(pBeneficiario)+" - "+Str(pLocal)+" - "+Str(pRegime)+" - "+Str(pCondicao)+" _ "+Str(pTipo)+" - "+Str(pObjetivo)+" - "+Str(pFinalidade)+" - "+Str(pRecebedor)+" - "+Str(pConvenio)+" - "+Str(pCodpagto)+" - "+Str(pXTHM)+" - "+Str(pEvento)+" - "+Str(pGrau)
 ' CancelDescription =  pOcorrencias
 ' CanContinue = False
 'Set Interface = CreateBennerObject("BSPRE010.Rotinas")
 'Interface.Processar(CurrentSystem,pBeneficiario, pLocal,pRegime,pCondicao,pTipo,pObjetivo,pFinalidade,pRecebedor,pLocalExecucao,pMunicipio,pEstado,pConvenio,pCodpagto,pXTHM,pAcomodacao,pEvento,pGrau,pData,pValorEvento,pValorPF,pQuantidade,pOcorrencias,pGrauEHMatMed)]


	Set SP = NewStoredProc
	SP.AutoMode = True
	SP.Name = "BSPRE_CALCULAPRECOTOTAL"

	SP.AddParam("p_Beneficiario",ptInput)
	SP.AddParam("p_LocalAtendimento",ptInput)
	SP.AddParam("p_RegimeAtendimento",ptInput)
	SP.AddParam("p_CondicaoAtendimento",ptInput)
	SP.AddParam("p_TipoTratamento",ptInput)
	SP.AddParam("p_ObjetivoTratamento",ptInput)
	SP.AddParam("p_FinalidadeAtendimento",ptInput)
	SP.AddParam("p_Recebedor",ptInput)
	SP.AddParam("p_LocalExecucao",ptInput)
	SP.AddParam("p_Municipio",ptInput)
	SP.AddParam("p_Estado",ptInput)
	SP.AddParam("p_Convenio",ptInput)
	SP.AddParam("p_Plano",ptInput)
	SP.AddParam("p_CodPagto",ptInput)
	SP.AddParam("p_XTHM",ptInput)
	SP.AddParam("p_Acomodacao",ptInput)
	SP.AddParam("p_Evento",ptInput)
	SP.AddParam("p_Grau",ptInput)
	SP.AddParam("p_TecnicaCirurgica",ptInput)
	SP.AddParam("p_Data",ptInput)
	SP.AddParam("p_Quantidade",ptInput)
	SP.AddParam("p_Chave",ptInput)
	SP.AddParam("p_Usuario",ptInput)
	SP.AddParam("p_LocalChamada",ptInput)
	SP.AddParam("p_DescontoCascata",ptInput)


	SP.AddParam("p_ValorEvento",ptOutput)
	sp.ParamByName("p_ValorEvento").DataType = ftFloat

	SP.AddParam("p_ValorPF",ptOutput)
	sp.ParamByName("p_ValorPF").DataType = ftFloat

	SP.AddParam("p_GrauEHMatMed",ptOutput)
	sp.ParamByName("p_GrauEHMatMed").DataType = ftString

    SP.ParamByName("p_Beneficiario").AsInteger        	= pBeneficiario
    SP.ParamByName("p_LocalAtendimento").AsInteger    	= pLocal
    SP.ParamByName("p_RegimeAtendimento").AsInteger   	= pRegime
    SP.ParamByName("p_CondicaoAtendimento").AsInteger 	= pCondicao
    SP.ParamByName("p_TipoTratamento").AsInteger      	= pTipo
    SP.ParamByName("p_ObjetivoTratamento").AsInteger    = pObjetivo
    SP.ParamByName("p_FinalidadeAtendimento").AsInteger = pFinalidade
    SP.ParamByName("p_Recebedor").AsInteger             = pRecebedor
    SP.ParamByName("p_LocalExecucao").AsInteger         = pLocalExecucao
    SP.ParamByName("p_Municipio").AsInteger             = pMunicipio
    SP.ParamByName("p_Estado").AsInteger          		= pEstado
    SP.ParamByName("p_Convenio").AsInteger          	= pConvenio
    SP.ParamByName("p_Plano").AsInteger          		= pPLano
    SP.ParamByName("p_CodPagto").AsInteger          	= pCodpagto
    SP.ParamByName("p_XTHM").AsInteger          		= pXTHM
    SP.ParamByName("p_Acomodacao").AsInteger          	= pAcomodacao
    SP.ParamByName("p_Evento").AsInteger          		= pEvento
    SP.ParamByName("p_Grau").AsInteger          		= pGrau
    SP.ParamByName("p_TecnicaCirurgica").AsString		= pTecnicaCirurgica
    SP.ParamByName("p_Data").AsDateTime          		= pData 'DateValue("30/10/2006")
    SP.ParamByName("p_Quantidade").AsInteger          	= pQuantidade
    SP.ParamByName("p_Chave").AsInteger          		= p_Beneficiario
    SP.ParamByName("p_Usuario").AsInteger          		= CurrentUser
    SP.ParamByName("p_LocalChamada").AsString = "W"
    SP.ParamByName("p_DescontoCascata").AsFloat = 0
    SP.ExecProc

   'If Not VisibleMode Then
   '  pOcorrencias = "pBeneficiario = "+Str(pBeneficiario)+" - "+"pLocalAtendimento = "+Str(pLocal)+ " - "+"pRegimeAtendimento = "+Str(pRegime)+" - "+"pCondicaoAtendimento = "+Str(pCondicao)+" - "+"pTipoTratamento = "+Str(pTipo)+" - "+"pObjetivoTratamento = "+Str(pObjetivo)+" - "+"pFinalidadeAtendimento = "+Str(pFinalidade)+" - "+"pRecebedor = "+Str(pRecebedor)+" - "+"pLocalExecucao = "+Str(pLocalExecucao)+" - "+"pMunicipio = "+Str(pMunicipio)+" - "+"pEstado = "+Str(pEstado)+" - "+"pConvenio = "+Str(pConvenio)+ _
   '                 " - pPLano = "+Str(pPLano)+" - "+"pCodpagto = "+Str(pCodpagto)+" - "+"pXTHM = "+Str(pXTHM)+" - "+"pAcomodacao = "+Str(pAcomodacao)+" - "+"pEvento = "+Str(pEvento)+" - "+"pGrau = "+Str(pGrau)+" - "+"pData = "+Str(pData)+" - "+"pQuantidade = "+Str(pQuantidade)+" - "+" CurrentUser = "+Str(CurrentUser)
   ' CancelDescription =  pOcorrencias
   '  CanContinue = False
   'End If
 '----------------------------------------------------------------------
  CurrentQuery.FieldByName("VALORPF").AsFloat  = SP.ParamByName("p_ValorPF").AsFloat
  pValorPF = SP.ParamByName("p_ValorPF").AsFloat
  pGrauEHMatMed = SP.ParamByName("p_GrauEHMatMed").AsString
  pOcorrencias = ""
  CurrentQuery.FieldByName("VALOR").AsFloat    = SP.ParamByName("p_ValorEvento").AsFloat
  pValorEvento =  SP.ParamByName("p_ValorEvento").AsFloat
 '----------------------------------------------------------------------
 If pOcorrencias = "" Then

    If pGrauEHMatMed = "S" Then
      pOcorrencias = "Neste preço não estão inclusos materiais e medicamentos."
    Else
      pOcorrencias = ""
    End If
  End If

  CurrentQuery.FieldByName("VALOR").AsFloat = pValorEvento
  CurrentQuery.FieldByName("VALORPF").AsFloat = pValorPF

  InfoDescription = pOcorrencias
 '--------------------------------------------
 'If Not VisibleMode Then
 '  CancelDescription =  pOcorrencias
 '  CanContinue = False
 'End If
 '--------------------------------------------
 'Set Interface = Nothing
  SP.AutoMode = False
End Sub

Public Sub TABLE_ExternalValidate(CanContinue As Boolean, ByVal Param As String)

End Sub

Public Sub TABLE_NewRecord()
  Dim  SQL            As Object
  Dim  qConvenio      As Object
  Dim  qAtendim       As Object
  Dim  qBeneficiario  As Object
  Dim  vUsuario       As Long


  Set SQL       = NewQuery
  Set qConvenio = NewQuery
  Set qAtendim  = NewQuery

  '------------------ BENEFICIARIO -----------------------------------
  If WebVisionCode = "2" Then
    vUsuario = CurrentUser
    Set qBeneficiario = NewQuery
	qBeneficiario.Add("SELECT A.HANDLE                                                                               ")
	qBeneficiario.Add("  FROM SAM_BENEFICIARIO A                                                                     ")
	qBeneficiario.Add(" WHERE A.FAMILIA IN (SELECT B.FAMILIA                                                         ")
	qBeneficiario.Add("                       FROM SAM_BENEFICIARIO             B                                    ")
	qBeneficiario.Add("                       JOIN SAM_MATRICULA                M  ON (B.MATRICULA = M.HANDLE)       ")
	qBeneficiario.Add("                       JOIN Z_GRUPOUSUARIOS_BENEFICIARIO ZB ON (ZB.MATRICULAUNICA = M.HANDLE) ")
	qBeneficiario.Add("                      WHERE ZB.USUARIO = :USUARIO)                                            ")
	qBeneficiario.Add("   AND A.EHTITULAR = 'S'                                                                      ")
	qBeneficiario.ParamByName("USUARIO").Value = vUsuario
	qBeneficiario.Active = True
	gBeneficiario = qBeneficiario.FieldByName("HANDLE").AsInteger
	qBeneficiario.Active = False
	CurrentQuery.FieldByName("BENEFICIARIO").Value = gBeneficiario
  End If
  '------------------XTHM E CODIGO DE PAGAMANTO ----------------------
  SQL.Add("SELECT CODIGOPAGTO,CODIGOXTHM FROM SAM_PARAMETROSPROCCONTAS")
  SQL.Active = True
  gXTHM = SQL.FieldByName("CODIGOXTHM").AsInteger
  gCodpagto = SQL.FieldByName("CODIGOPAGTO").AsInteger
  SQL.Active = False

  CurrentQuery.FieldByName("XTHM").Value = gXTHM
  CurrentQuery.FieldByName("CODPAGAMENTO").Value = gCodpagto
  '-------------------- CONVENIO --------------------------------------
  If WebVisionCode = "2" Then
	  qConvenio.Add("SELECT B.PLANO,                                       ")
	  qConvenio.Add("       B.CONVENIO                                     ")
	  qConvenio.Add("  FROM SAM_BENEFICIARIO  A                            ")
	  qConvenio.Add("  JOIN SAM_CONTRATO      B ON (A.CONTRATO = B.HANDLE) ")
	  qConvenio.Add(" WHERE A.HANDLE = :BENEFICIARIO                       ")
	  qConvenio.ParamByName("BENEFICIARIO").Value = gBeneficiario
	  qConvenio.Active = True
	  gConvenio = qConvenio.FieldByName("CONVENIO").AsInteger
	  gPlano    = qConvenio.FieldByName("PLANO").AsInteger
  Else
	  qConvenio.Add("SELECT HANDLE FROM SAM_CONVENIO WHERE HANDLE = (SELECT MIN(HANDLE) FROM SAM_CONVENIO)")
	  qConvenio.Active = True
	  gConvenio = qConvenio.FieldByName("HANDLE").AsInteger
	  CurrentQuery.FieldByName("CONVENIO").Value = gConvenio
	  qConvenio.Active = False
  End If
  CurrentQuery.FieldByName("CONVENIO").Value = gConvenio
  qConvenio.Active = False
  '------------------ CARACTERISTICAS DE ATENDIMENTO ------------------

  qAtendim.Add("SELECT FINALIDADEATENDIMENTO,    ")
  qAtendim.Add("       LOCALATENDIMENTO,         ")
  qAtendim.Add("       REGIMEATENDIMENTO,        ")
  qAtendim.Add("       CONDICAOATENDIMENTO,      ")
  qAtendim.Add("       OBJETIVOTRATAMENTO,       ")
  qAtendim.Add("       TIPOTRATAMENTO            ")
  qAtendim.Add("  FROM SAM_PARAMETROSPRESTADOR   ")

  qAtendim.Active = True
  gLocal      = qAtendim.FieldByName("LOCALATENDIMENTO").AsInteger
  gRegime     = qAtendim.FieldByName("REGIMEATENDIMENTO").AsInteger
  gCondicao   = qAtendim.FieldByName("CONDICAOATENDIMENTO").AsInteger
  gTipo       = qAtendim.FieldByName("TIPOTRATAMENTO").AsInteger
  gObjetivo   = qAtendim.FieldByName("OBJETIVOTRATAMENTO").AsInteger
  gFinalidade = qAtendim.FieldByName("FINALIDADEATENDIMENTO").AsInteger

  CurrentQuery.FieldByName("LOCALATENDIMENTO").Value = gLocal
  CurrentQuery.FieldByName("REGIMEATENDIMENTO").Value = gRegime
  CurrentQuery.FieldByName("CONDATENDIMENTO").Value = gCondicao
  CurrentQuery.FieldByName("TIPOTRATAMENTO").Value = gTipo
  CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").Value = gObjetivo
  CurrentQuery.FieldByName("FINALIDADEATENDIMENTO").Value = gFinalidade
  qAtendim.Active = False

  SQL.Clear
  SQL.Add("SELECT PRESTADOR FROM Z_GRUPOUSUARIOS_PRESTADOR WHERE USUARIO = :USUARIO")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser
  SQL.Active = True

  If Not SQL.FieldByName("PRESTADOR").IsNull Then
    CurrentQuery.FieldByName("RECEBEDOR").AsInteger = SQL.FieldByName("PRESTADOR").AsInteger
    CurrentQuery.FieldByName("LOCALEXECUCAO").AsInteger = SQL.FieldByName("PRESTADOR").AsInteger
  End If
  Set SQL = Nothing

End Sub
