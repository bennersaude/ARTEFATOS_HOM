'HASH: F6A923A67D6C5ACE3B59DA505C2DACD3

Option Explicit

Public Sub AlteraSituacaoSolicitacao(pHandle As Long, pSituacao As String)
  Dim sql As Object
  Set sql = NewQuery

  If Not InTransaction Then
    StartTransaction
  End If

  sql.Clear
  sql.Add("UPDATE WEB_CADASTRO SET SITUACAO = :situacao, USUARIOCANCELAMENTO = :USUARIO, DATACANCELAMENTO = :DATA WHERE HANDLE = :HANDLE")
  sql.ParamByName("USUARIO").AsInteger = CurrentUser
  sql.ParamByName("HANDLE").AsInteger = pHandle
  sql.ParamByName("DATA").AsDateTime = ServerNow
  sql.ParamByName("SITUACAO").AsString = pSituacao
  sql.ExecSQL

  If InTransaction Then
    Commit
  End If


  Set sql = Nothing

  RefreshNodesWithTable("WEB_CADASTRO")

End Sub
Public Function ChecarIdadeMaximaDependente(pBeneficiario As Long, SituacaoRh As Long, ByRef Idade As Long, ByRef IdadeMaxima As Long) As Boolean
  Dim sql As Object
  Set sql = NewQuery


  sql.Clear
  sql.Add("SELECT M.DATANASCIMENTO, A.HANDLE FROM SAM_CONTRATO_TPDEP A")
  sql.Add("  JOIN SAM_BENEFICIARIO B ON (B.TIPODEPENDENTE = A.HANDLE)")
  sql.Add("  JOIN SAM_MATRICULA    M ON (B.MATRICULA = M.HANDLE) ")
  sql.Add(" WHERE B.HANDLE = :HANDLE ")
  sql.ParamByName("HANDLE").AsInteger = pBeneficiario
  sql.Active = True

  Dim TPDepBenef As Long
  Dim vDataNascimento As Date

  vDataNascimento = sql.FieldByName("DATANASCIMENTO").AsDateTime

  TPDepBenef = sql.FieldByName("HANDLE").AsInteger


  sql.Clear
  sql.Add("SELECT B.DEPENDENTEDESTINO")
  sql.Add("  FROM SAM_CONTRATO_TPDEP A")
  sql.Add("  JOIN SAM_SITUACAORH_TPDEP     B ON (B.DEPENDENTEORIGEM = A.HANDLE) ")
  sql.Add(" WHERE B.SITUACAORH = :HANDLE")
  sql.Add("  AND B.DEPENDENTEORIGEM = :DEP")
  sql.ParamByName("HANDLE").AsInteger = SituacaoRh
  sql.ParamByName("DEP").AsInteger = TPDepBenef

  sql.Active = True

  Dim TpDepDestino As Long
  TpDepDestino = sql.FieldByName("DEPENDENTEDESTINO").AsInteger

  sql.Clear
  sql.Add("SELECT A.IDADEMAXIMA FROM SAM_CONTRATO_TPDEP A")
  sql.Add(" WHERE a.HANDLE = :HANDLE ")
  sql.ParamByName("HANDLE").AsInteger = TpDepDestino
  sql.Active = True

  If sql.FieldByName("IDADEMAXIMA").IsNull Then
    ChecarIdadeMaximaDependente = True
  Else
    Idade = DateDiff("yyyy", vDataNascimento,ServerDate)

    IdadeMaxima = sql.FieldByName("IDADEMAXIMA").AsInteger

    If Idade > sql.FieldByName("IDADEMAXIMA").AsInteger Then
      ChecarIdadeMaximaDependente = False
    Else
      ChecarIdadeMaximaDependente = True
    End If


  End If
End Function


Public Function BeneficiarioPodeAtivar(pBeneficiario As Long, pDataBase As Date, ByRef pDataMinima As Date, pProcessoAcao As Integer) As Boolean
  If (Not VisibleMode) And ((pProcessoAcao = 10) Or (pProcessoAcao = 15)) Then
    Dim sql As Object
    Set sql = NewQuery

    If pProcessoAcao = 10 Then
      sql.Clear
      sql.Add("SELECT B.DATACANCELAMENTO, C.PRAZOMINIMOREATIVACAO")
      sql.Add("  FROM SAM_BENEFICIARIO  B")
      sql.Add("  JOIN SAM_MATRICULA     M ON (B.MATRICULA = M.HANDLE) ")
      sql.Add("  JOIN Z_GRUPOUSUARIOS_BENEFICIARIO ZB ON (ZB.MATRICULAUNICA = M.HANDLE)")
      sql.Add("  JOIN SAM_MOTIVOCANCELAMENTO C ON (B.MOTIVOCANCELAMENTO = C.HANDLE)")
      sql.Add(" WHERE DATACANCELAMENTO IS NOT NULL ")
      sql.Add(" AND EHTITULAR = 'S'")
      sql.Add("  AND ZB.USUARIO = :USUARIO")
      sql.Add(" ORDER BY DATACANCELAMENTO DESC")
      sql.ParamByName("USUARIO").AsInteger = CurrentUser
      sql.Active = True
    End If

    If pProcessoAcao = 15 Then
'      sql.Add(" AND EHTITULAR = 'N'")
'      sql.Clear
      'sql.Add("SELECT B.DATACANCELAMENTO, C.PRAZOMINIMOREATIVACAO")
'      sql.Add("  FROM SAM_BENEFICIARIO  B")
      'sql.Add("  JOIN SAM_MOTIVOCANCELAMENTO C ON (B.MOTIVOCANCELAMENTO = C.HANDLE)")
'      sql.Add(" WHERE B.HANDLE = :HANDLE AND DATACANCELAMENTO IS NOT NULL ")
'      sql.Add(" AND EHTITULAR = 'N'")
'      sql.ParamByName("HANDLE").AsInteger = pBeneficiario
      'sql.Active = True
      sql.Clear
      sql.Add("SELECT B.DATACANCELAMENTO, C.PRAZOMINIMOREATIVACAO")
      sql.Add("  FROM SAM_BENEFICIARIO B ")
      sql.Add("  JOIN SAM_MATRICULA    M ON (B.MATRICULA = M.HANDLE)")
      sql.Add("  JOIN SAM_MOTIVOCANCELAMENTO C ON (B.MOTIVOCANCELAMENTO = C.HANDLE)")
      sql.Add(" WHERE B.FAMILIA IN (SELECT FAMILIA")
      sql.Add("                      FROM SAM_BENEFICIARIO   B")
      sql.Add("                      JOIN SAM_MATRICULA      M ON (B.MATRICULA = M.HANDLE)")
      sql.Add("                      JOIN Z_GRUPOUSUARIOS_BENEFICIARIO Z ON (Z.MATRICULAUNICA = M.HANDLE)")
      sql.Add("                     WHERE Z.USUARIO = :USUARIO)")
      sql.Add("   AND B.MATRICULA = (SELECT MATRICULA FROM SAM_BENEFICIARIO WHERE HANDLE = :BENEFICIARIO)")
      sql.Add("   AND B.DATACANCELAMENTO IS NOT NULL")
      sql.Add("ORDER BY B.DATACANCELAMENTO DESC")
      sql.ParamByName("USUARIO").AsInteger = CurrentUser
      sql.ParamByName("BENEFICIARIO").AsInteger = pBeneficiario
      sql.Active= True


    End If

    If sql.EOF Then
      'Deve checar a idade do tipo de dependente
      BeneficiarioPodeAtivar = True

    Else
      Dim vDataMinimaReativacao As Date

      vDataMinimaReativacao = DateAdd("d",sql.FieldByName("PRAZOMINIMOREATIVACAO").AsInteger,sql.FieldByName("DATACANCELAMENTO").AsDateTime)

      pDataMinima = vDataMinimaReativacao

      If vDataMinimaReativacao >= pDataBase Then
        BeneficiarioPodeAtivar = False
      Else
        BeneficiarioPodeAtivar = True
      End If
    End If
    Set sql = Nothing
  Else
    BeneficiarioPodeAtivar = True
  End If

End Function


Public Function ProcessarSolicitacao(pHandle As Long, pBeneficiario As Long, pAcao As Long) As String
  Dim vDLL As Object
  Dim sql As Object
  Dim vResultado As String
  Dim vFamilia As Integer
  Dim vSituacaoRH As Long
  Dim vExisteDocumentoPendente As Boolean
  Dim vContratoDestino As Long
  Dim vContratoOrigem As Long
  Dim vDocumentosPendentes As String
  Dim vCodigoProcessoAcao As Integer
  Dim vSituacaoRHAnterior As Long
  Dim vMatriculaUnica As Long


  On Error GoTo Erro

  Set sql = NewQuery
  sql.Clear
  sql.Add("SELECT FAMILIA, SITUACAORH, MATRICULA FROM SAM_BENEFICIARIO WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = pBeneficiario
  sql.Active = True

  vFamilia = sql.FieldByName("FAMILIA").AsInteger
  vSituacaoRHAnterior = sql.FieldByName("SITUACAORH").AsInteger
  vMatriculaUnica = sql.FieldByName("MATRICULA").AsInteger


  sql.Clear
  sql.Add("SELECT A.SITUACAORH, B.CONTRATOMIGRACAO, C.CODIGO")
  sql.Add("  FROM WEB_CADASTROACAO A")
  sql.Add("  JOIN WEB_CADASTROPROCESSO C ON (A.WEBCADASTROPROCESSO = C.HANDLE)")
  sql.Add("  JOIN SAM_SITUACAORH   B ON (A.SITUACAORH = B.HANDLE)")
  sql.Add(" WHERE A.HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = pAcao
  sql.Active = True

  vSituacaoRH = sql.FieldByName("SITUACAORH").AsInteger
  vContratoDestino = sql.FieldByName("CONTRATOMIGRACAO").AsInteger
  vCodigoProcessoAcao = sql.FieldByName("CODIGO").AsInteger

  Set vDLL = CreateBennerObject("BSBEN005.Rotinas")

  Dim vDataStr As String
  vDataStr = CurrentQuery.FieldByName("DATAOPERACAO").AsString


  If Not VisibleMode Then 'Somente checa na web. Caso contrário não checa, pois o usuário processará automaticamente depois

      'Verificar se nao atingiu a idade maxima
      Dim vIdade As Long
      Dim vIdadeMaxima As Long


      If Not ChecarIdadeMaximaDependente(pBeneficiario,vSituacaoRH, vIdade, vIdadeMaxima ) Then
        ProcessarSolicitacao = "Operação não realizada. Idade máxima do dependente atingida. Idade máxima permitida: " + Str(vIdadeMaxima) + ". Idade do beneficiário: " + Str(vIdade)
        If Not InTransaction Then
          StartTransaction
        End If

        sql.Clear
        sql.Add("UPDATE WEB_CADASTRO SET OCORRENCIAS = :TEXTO, SITUACAO = 'C' WHERE HANDLE = :HANDLE")
        sql.ParamByName("TEXTO").AsString = "Operação não realizada. Idade máxima do dependente atingida. Idade máxima permitida: " + Str(vIdadeMaxima) + ". Idade do beneficiário: " + Str(vIdade)
        sql.ParamByName("HANDLE").AsInteger = pHandle
        sql.ExecSQL

        Commit

        Exit Function
      End If


  vExisteDocumentoPendente = vDLL.VerificaDocumentosPendentes (CurrentSystem, _
    vSituacaoRH, _
    pBeneficiario, _
    vContratoDestino, _
    CurrentQuery.FieldByName("DATAOPERACAO").AsString, _
    vDocumentosPendentes)
  Else
    vExisteDocumentoPendente = False
  End If


  If Not vExisteDocumentoPendente Then
    Dim vDataMinimaReativacao As Date

    If BeneficiarioPodeAtivar(pBeneficiario,CurrentQuery.FieldByName("DATAOPERACAO").AsDateTime, vDataMinimaReativacao, vCodigoProcessoAcao) Then


      vResultado = vDLL.SituacaoRHMigracao(CurrentSystem, pBeneficiario, _
      vDataStr, vFamilia, vSituacaoRH, False, False)
      If Len(vResultado) > 0 Then
         ProcessarSolicitacao = vResultado
      Else
        ProcessarSolicitacao = "Solicitação processada com sucesso."
        If Not InTransaction Then
          StartTransaction
        End If

        ' Atualiza tanto o beneficiário origem quanto o beneficiário destino, com a situacao rh anterior.

        sql.Clear
		sql.Add("SELECT A.TABELASITUACAORH, TABTIPOSITUACAO FROM SAM_SITUACAORH A")
		sql.Add("WHERE A.HANDLE = :HANDLE")
		sql.ParamByName("HANDLE").AsInteger = vSituacaoRHAnterior
		sql.Active = True

		Dim tabelaSituacaorh As Long
		Dim tabSituacao As Integer
		Dim vNovaSituacaoRHDestino As Long

		tabelaSituacaorh = sql.FieldByName("TABELASITUACAORH").AsInteger
		tabSituacao = sql.FieldByName("TABTIPOSITUACAO").AsInteger

		sql.Clear
		sql.Add("SELECT HANDLE FROM SAM_SITUACAORH WHERE CONTRATO = :CONTRATO")
		sql.Add("AND TABELASITUACAORH = :TABELA AND TABTIPOSITUACAO = :TAB")
		sql.ParamByName("CONTRATO").AsInteger = vContratoDestino
		sql.ParamByName("TAB").AsInteger = tabSituacao
		sql.ParamByName("TABELA").AsInteger = tabelaSituacaorh
		sql.Active = True

		vNovaSituacaoRHDestino = sql.FieldByName("HANDLE").AsInteger

 		If vNovaSituacaoRHDestino  > 0 Then
 		  sql.Clear
          sql.Add("UPDATE SAM_BENEFICIARIO SET SITUACAORH = :SITUACAOANTERIOR WHERE MATRICULA = :HANDLE AND CONTRATO = :CONTRATODESTINO")
          sql.ParamByName("SITUACAOANTERIOR").AsInteger = vNovaSituacaoRHDestino
        sql.ParamByName("HANDLE").AsInteger = vMatriculaUnica
          sql.ParamByName("CONTRATODESTINO").AsInteger = vContratoDestino
        sql.ExecSQL
 		End If


        If vSituacaoRHAnterior > 0 Then
          sql.Clear
          sql.Add("UPDATE SAM_BENEFICIARIO SET SITUACAORH = :SITUACAOANTERIOR WHERE HANDLE = :HANDLE")
          sql.ParamByName("SITUACAOANTERIOR").AsInteger = vSituacaoRHAnterior
          sql.ParamByName("HANDLE").AsInteger = pBeneficiario
          sql.ExecSQL
        End If

        sql.Clear
        sql.Add("UPDATE WEB_CADASTRO SET SITUACAO = 'P', USUARIOPROCESSAMENTO = :USUARIO, DATAPROCESSAMENTO = :DATA WHERE HANDLE = :HANDLE")
        sql.ParamByName("HANDLE").AsInteger = pHandle
        sql.ParamByName("USUARIO").AsInteger = CurrentUser
        sql.ParamByName("DATA").AsDateTime = ServerNow
        sql.ExecSQL
        Commit
      End If
    Else
      vDataMinimaReativacao = DateAdd("d",1,vDataMinimaReativacao)
      ProcessarSolicitacao = "Operação não realizada. Prazo mínimo para reativação no plano não atingido. Data mínima para reativação: " +  Trim(Str(Format(Day(vDataMinimaReativacao),"00")))+"/"+Trim(Str(Format(Month(vDataMinimaReativacao),"00")))+"/"+Trim(Str(Year(vDataMinimaReativacao)))
      If Not InTransaction Then
        StartTransaction
      End If

      sql.Clear
      sql.Add("UPDATE WEB_CADASTRO SET OCORRENCIAS = :TEXTO, SITUACAO = 'C' WHERE HANDLE = :HANDLE")
      sql.ParamByName("TEXTO").AsString = "Prazo mínimo para reativação no plano não atingido. Data mínima para reativação: " + Trim(Str(Format(Day(vDataMinimaReativacao),"00")))+"/"+Trim(Str(Format(Month(vDataMinimaReativacao),"00")))+"/"+Trim(Str(Year(vDataMinimaReativacao)))
      sql.ParamByName("HANDLE").AsInteger = pHandle
      sql.ExecSQL

      Commit

    End If
  Else
    ProcessarSolicitacao = "Existem documentos pendentes: "+ vDocumentosPendentes
    If Not InTransaction Then
      StartTransaction
    End If
    If Not InTransaction Then
      StartTransaction
    End If

    sql.Clear
    sql.Add("UPDATE WEB_CADASTRO SET OCORRENCIAS = :TEXTO WHERE HANDLE = :HANDLE")
    sql.ParamByName("TEXTO").AsString = "Existem documentos pendentes: "+ vDocumentosPendentes
    sql.ParamByName("HANDLE").AsInteger = pHandle
    sql.ExecSQL

    If InTransaction Then
      Commit
    End If

  End If

  Set vDLL = Nothing
  Set sql = Nothing

  Exit Function

  Erro:
    ProcessarSolicitacao = Err.Description
    Set vDLL = Nothing
    Set sql = Nothing


End Function



Public Sub BOTAOCANCELAR_OnClick()
  AlteraSituacaoSolicitacao CurrentQuery.FieldByName("HANDLE").AsInteger,"C"
  RefreshNodesWithTable("WEB_CADASTRO")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  MsgBox ProcessarSolicitacao(CurrentQuery.FieldByName("HANDLE").AsInteger,CurrentQuery.FieldByName("BENEFICIARIO").AsInteger,CurrentQuery.FieldByName("ACAO").AsInteger)
  RefreshNodesWithTable("WEB_CADASTRO")
End Sub

Public Sub BOTAOREATIVARSOLICITACAO_OnClick()
  AlteraSituacaoSolicitacao CurrentQuery.FieldByName("HANDLE").AsInteger,"A"
  RefreshNodesWithTable("WEB_CADASTRO")
End Sub

Public Sub TABLE_AfterInsert()
  Dim sql As Object
  Set sql = NewQuery
  'Texto comum a todas as visões
  CurrentQuery.FieldByName("USUARIOSOLICITACAO").AsInteger = CurrentUser
  CurrentQuery.FieldByName("DATASOLICITACAO").AsDateTime = ServerNow
  CurrentQuery.FieldByName("DATAOPERACAO").AsDateTime = ServerDate+1

  sql.Clear
  sql.Add("SELECT A.HANDLE, MENSAGEM, B.CODIGO")
  sql.Add("  FROM WEB_CADASTROACAO   A")
  sql.Add("  JOIN WEB_CADASTROPROCESSO B ON (A.WEBCADASTROPROCESSO = B.HANDLE)")
  sql.Add(" WHERE IDENTIFICADORVISAO = :ID")
  sql.ParamByName("ID").AsString = WebVisionCode
  sql.Active = True

  CurrentQuery.FieldByName("ACAO").AsInteger = sql.FieldByName("HANDLE").AsInteger
  CurrentQuery.FieldByName("OCORRENCIAS").AsString = sql.FieldByName("MENSAGEM").AsString

  'Códigos específicos
  'If (WebVisionCode = "1") Or (WebVisionCode = "2") Or (WebVisionCode = "3") Then
    WebVisionCode1("NR", 0, sql.FieldByName("CODIGO").AsInteger)
  'End If

  Set sql = Nothing

End Sub

Public Sub TABLE_AfterPost()
  Dim vResultado As String

  If InTransaction Then
    Commit 'O commit está aqui, porque mesmo que a solicitação nao seja processada, ela deve ser gravada para que o usuário do sistema possa tomar alguma ação
  End If

  vResultado = ProcessarSolicitacao(CurrentQuery.FieldByName("HANDLE").AsInteger,CurrentQuery.FieldByName("BENEFICIARIO").AsInteger,CurrentQuery.FieldByName("ACAO").AsInteger)

  InfoDescription = "Anote o número de seu protocolo: " + CurrentQuery.FieldByName("PROTOCOLO").AsString + Chr(13) + Chr(10) + vResultado

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  'Comum a todas as visões
  Dim vProtocolo As Long
  NewCounter("WEB_PROTOCOLO",Year(ServerDate), 1,vProtocolo)

  CurrentQuery.FieldByName("PROTOCOLO").AsString = Trim(Str(Year(ServerDate))) + Trim(Str(vProtocolo))
  'O Protocolo sempre será o ano corrente + protocolo. O contador é reiniciado por ano

  CurrentQuery.FieldByName("OCORRENCIAS").Clear

  If CurrentQuery.FieldByName("DATAOPERACAO").AsDateTime <= ServerDate Then
    CancelDescription = "A data somente poder ser maior que a data atual"
    CanContinue = False
  End If



End Sub



Public Function WebVisionCode1(pModo As String , ByVal pBeneficiario As Integer, pCodigoProcesso As Integer) As Boolean
	Dim sql As Object
	Dim vMensagem As String 'Variável utilizada para emitir mensagem para quando não existir registro para cadastro

	Set sql = NewQuery


    If (pModo = "BP") Then
 	  'SQL.Clear
      'SQL.Add("SELECT A.HANDLE")
      'SQL.Add("  FROM WEB_CADASTROACAO   A")
      'SQL.Add("  JOIN SAM_SITUACAORH     B ON    (A.SITUACAORH = B.HANDLE)")
      'SQL.Add(" WHERE B.TABTIPOSITUACAO  = 2")
      'SQL.Add("   AND B.CONTRATO = (SELECT C.CONTRATO")
      'SQL.Add("                       FROM SAM_BENEFICIARIO C")
      'SQL.Add("                      WHERE C.HANDLE = :BENEFICIARIO)")
      'SQL.ParamByName("BENEFICIARIO").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger 'pBeneficiario
      'SQL.Active = True
      'CurrentQuery.FieldByName("ACAO").AsInteger = SQL.FieldByName("HANDLE").AsInteger
      'Ação será informada pelo usuário ou gravada pela própria visão.

    ElseIf (pModo = "NR") Then

      Dim vSelectEspecial As String
      'Checa se o beneficiário que é da família do usuário que logou-se no sistema segundo o webvision code, que determina o contrato da operaçao

      If (pCodigoProcesso = 10)  Then
        'Inclusão de beneficiário titular via migração
        vSelectEspecial = "      (A.FAMILIA IN (SELECT B.FAMILIA " + Chr(13)
        vSelectEspecial = vSelectEspecial +   "FROM SAM_BENEFICIARIO B" + Chr(13)
        vSelectEspecial = vSelectEspecial +   "JOIN SAM_MATRICULA M ON (B.MATRICULA = M.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial +   "JOIN Z_GRUPOUSUARIOS_BENEFICIARIO ZB ON (ZB.MATRICULAUNICA = M.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial + " WHERE ZB.USUARIO = "+ Str(CurrentUser) + ")" + Chr(13)
        vSelectEspecial = vSelectEspecial + "   And A.CONTRATO = (SELECT CONTRATO" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                     FROM SAM_SITUACAORH   A" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                     JOIN WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                    WHERE IDENTIFICADORVISAO = '" + WebVisionCode + "')"
        vSelectEspecial = vSelectEspecial + " )"

        vSelectEspecial = vSelectEspecial + " AND A.DATACANCELAMENTO IS NULL"
        vSelectEspecial = vSelectEspecial + " AND EHTITULAR = 'S'"

        vMensagem = "Não existe titular para ser incluído!"
      End If


      If (pCodigoProcesso = 15) Then
        'Inclusão de beneficiário dependente via migração

        'vSelectEspecial = vSelectEspecial + " AND EHTITULAR = 'N'"
        'Verifica se existe um beneficiário titular ativo na família do contrato de destino
        'vSelectEspecial = vSelectEspecial + "  And EXISTS (Select 1 "
        'vSelectEspecial = vSelectEspecial + "       FROM SAM_BENEFICIARIO X"
        'vSelectEspecial = vSelectEspecial + "      WHERE CONTRATO = (Select CONTRATOMIGRACAO FROM SAM_SITUACAORH A"
        'vSelectEspecial = vSelectEspecial + "                         Join WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE)"
        'vSelectEspecial = vSelectEspecial + "                        WHERE B.IDENTIFICADORVISAO = '"+WebVisionCode+"')"
        'vSelectEspecial = vSelectEspecial + "        And FAMILIA In (Select B.FAMILIA"
        'vSelectEspecial = vSelectEspecial + "                        FROM SAM_BENEFICIARIO B"
        'vSelectEspecial = vSelectEspecial + "                        Join SAM_MATRICULA M On (B.MATRICULA = M.HANDLE)"
        'vSelectEspecial = vSelectEspecial + "                        Join Z_GRUPOUSUARIOS_BENEFICIARIO ZB On (ZB.MATRICULAUNICA = M.HANDLE)"
        'vSelectEspecial = vSelectEspecial + "                       WHERE ZB.USUARIO = "+Str(CurrentUser) + ")"
        'vSelectEspecial = vSelectEspecial + "        And EHTITULAR = 'S'"
        'vSelectEspecial = vSelectEspecial + "        And DATACANCELAMENTO Is  Null)"

        'vSelectEspecial = vSelectEspecial + " A.FAMILIA IN (SELECT F.HANDLE FROM SAM_FAMILIA F WHERE  " + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " F.FAMILIA IN (SELECT F1.FAMILIA" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                       FROM SAM_BENEFICIARIO  B" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                       JOIN SAM_FAMILIA       F1 ON (B.FAMILIA = F1.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                      WHERE B.FAMILIA IN ( SELECT B.FAMILIA" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                                            FROM SAM_BENEFICIARIO B" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                                            JOIN SAM_MATRICULA M ON (B.MATRICULA = M.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                                            JOIN Z_GRUPOUSUARIOS_BENEFICIARIO ZB ON (ZB.MATRICULAUNICA = M.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                                           WHERE ZB.USUARIO = " +Str(CurrentUser) +")" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                      )       )   " + Chr(13) + Chr(10) ' segundo fecham é do primeiro familia
        'vSelectEspecial = vSelectEspecial + " AND A.CONTRATO = (SELECT CONTRATO" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                     FROM SAM_SITUACAORH   A" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                     JOIN WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                    WHERE B.IDENTIFICADORVISAO = '" + WebVisionCode + "'" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                  )" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " AND A.DATACANCELAMENTO IS NULL" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " AND EHTITULAR = 'N' " + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " And EXISTS(Select 1 FROM SAM_BENEFICIARIO X" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " Join SAM_FAMILIA FX On (X.FAMILIA = FX.Handle)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " WHERE X.CONTRATO = (Select CONTRATOMIGRACAO" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " FROM SAM_SITUACAORH   A" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " Join WEB_CADASTROACAO B On (B.SITUACAORH = A.Handle)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + " WHERE B.IDENTIFICADORVISAO = '" + WebVisionCode + "')" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "  And X.EHTITULAR = 'S' AND FX.FAMILIA = (SELECT Z.FAMILIA FROM SAM_FAMILIA Z WHERE HANDLE = A.FAMILIA) AND X.DATACANCELAMENTO IS NULL)" + Chr(13) + Chr(10)

        vSelectEspecial = vSelectEspecial + "A.Handle In (Select G.Handle "
        vSelectEspecial = vSelectEspecial + "               FROM SAM_BENEFICIARIO G"
        vSelectEspecial = vSelectEspecial + "               Join SAM_FAMILIA      F On (G.FAMILIA = F.Handle)"
        vSelectEspecial = vSelectEspecial + "              where F.Handle In ( Select X.FAMILIA"
        vSelectEspecial = vSelectEspecial + "                                    FROM SAM_BENEFICIARIO X"
        vSelectEspecial = vSelectEspecial + "                                    Join SAM_MATRICULA    N On (X.MATRICULA = N.Handle)"
        vSelectEspecial = vSelectEspecial + "                                    Join Z_GRUPOUSUARIOS_BENEFICIARIO ZB On (ZB.MATRICULAUNICA = N.Handle)"
        vSelectEspecial = vSelectEspecial + "                                    Join SAM_SITUACAORH   C On (X.CONTRATO   = C.CONTRATO)"
        vSelectEspecial = vSelectEspecial + "                                    Join WEB_CADASTROACAO D On (D.SITUACAORH = C.Handle)"
        vSelectEspecial = vSelectEspecial + "                                   WHERE ZB.USUARIO           =  " + Str(CurrentUser)
        vSelectEspecial = vSelectEspecial + "                                     And D.IDENTIFICADORVISAO = '" + WebVisionCode + "'
        vSelectEspecial = vSelectEspecial + "                                     And X.EHTITULAR = 'S'"
        vSelectEspecial = vSelectEspecial + "                                 )"
        vSelectEspecial = vSelectEspecial + "                And EHTITULAR = 'N' "
        vSelectEspecial = vSelectEspecial + "                And G.DATACANCELAMENTO Is Null"
        vSelectEspecial = vSelectEspecial + "                And EXISTS( Select  1"
        vSelectEspecial = vSelectEspecial + "                              FROM SAM_BENEFICIARIO X"
        vSelectEspecial = vSelectEspecial + "                              Join SAM_FAMILIA FX On (X.FAMILIA = FX.Handle)"
        vSelectEspecial = vSelectEspecial + "                             WHERE X.CONTRATO = (Select CONTRATOMIGRACAO"
        vSelectEspecial = vSelectEspecial + "                                                   FROM SAM_SITUACAORH   A"
        vSelectEspecial = vSelectEspecial + "                                                   Join WEB_CADASTROACAO B On (B.SITUACAORH = A.Handle)"
        vSelectEspecial = vSelectEspecial + "                                                  WHERE B.IDENTIFICADORVISAO = '" + WebVisionCode +"' )
        vSelectEspecial = vSelectEspecial + "                               And X.EHTITULAR = 'S' "
        vSelectEspecial = vSelectEspecial + "                               And X.DATACANCELAMENTO Is Null"
        vSelectEspecial = vSelectEspecial + "                          )"
        vSelectEspecial = vSelectEspecial + "            )"

        'vSelectEspecial = vSelectEspecial + ""
        'vSelectEspecial = vSelectEspecial + ""

          vMensagem = "Não existe dependente para ser incluído"

      End If

      If pCodigoProcesso = 20 Then
        vSelectEspecial = "      (A.FAMILIA IN (SELECT B.FAMILIA " + Chr(13)
        vSelectEspecial = vSelectEspecial +   "FROM SAM_BENEFICIARIO B" + Chr(13)
        vSelectEspecial = vSelectEspecial +   "JOIN SAM_MATRICULA M ON (B.MATRICULA = M.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial +   "JOIN Z_GRUPOUSUARIOS_BENEFICIARIO ZB ON (ZB.MATRICULAUNICA = M.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial + " WHERE ZB.USUARIO = "+ Str(CurrentUser) + ")" + Chr(13)
        vSelectEspecial = vSelectEspecial + "   And A.CONTRATO = (SELECT CONTRATO" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                     FROM SAM_SITUACAORH   A" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                     JOIN WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                    WHERE IDENTIFICADORVISAO = '" + WebVisionCode + "')"
        vSelectEspecial = vSelectEspecial + " )"

        vSelectEspecial = vSelectEspecial + " AND A.DATACANCELAMENTO IS NULL"
        vSelectEspecial = vSelectEspecial + " AND A.EHTITULAR = 'S' "

        vMensagem = "Não existe titular para ser cancelado!"

      End If

      If pCodigoProcesso = 25 Then
        'Cancelamento de beneficiário dependente
        'vSelectEspecial = vSelectEspecial + " AND EHTITULAR = 'N'"  + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "  AND EXISTS (SELECT 1 " + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "       FROM SAM_BENEFICIARIO X" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "      WHERE CONTRATO = (Select CONTRATOMIGRACAO FROM SAM_SITUACAORH A" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                         Join WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                        WHERE B.IDENTIFICADORVISAO = '"+WebVisionCode+"')" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "        And FAMILIA In (Select B.FAMILIA" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                        FROM SAM_BENEFICIARIO B" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                        Join SAM_MATRICULA M On (B.MATRICULA = M.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                        Join Z_GRUPOUSUARIOS_BENEFICIARIO ZB On (ZB.MATRICULAUNICA = M.HANDLE)" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                       WHERE ZB.USUARIO = "+Str(CurrentUser) + ")" + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "        And DATACANCELAMENTO Is NOT Null)"+ Chr(13) + Chr(10)

        vSelectEspecial = "      (A.FAMILIA IN (SELECT B.FAMILIA " + Chr(13)
        vSelectEspecial = vSelectEspecial +   "FROM SAM_BENEFICIARIO B" + Chr(13)
        vSelectEspecial = vSelectEspecial +   "JOIN SAM_MATRICULA M ON (B.MATRICULA = M.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial +   "JOIN Z_GRUPOUSUARIOS_BENEFICIARIO ZB ON (ZB.MATRICULAUNICA = M.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial + " WHERE ZB.USUARIO = "+ Str(CurrentUser) + ")" + Chr(13)
        vSelectEspecial = vSelectEspecial + "   And A.CONTRATO = (SELECT CONTRATO" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                     FROM SAM_SITUACAORH   A" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                     JOIN WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE)" + Chr(13)
        vSelectEspecial = vSelectEspecial + "                    WHERE IDENTIFICADORVISAO = '" + WebVisionCode + "')"
        vSelectEspecial = vSelectEspecial + " )"

        vSelectEspecial = vSelectEspecial + " AND A.DATACANCELAMENTO IS NULL"
        vSelectEspecial = vSelectEspecial + " AND A.EHTITULAR = 'N' "

		'vSelectEspecial = vSelectEspecial + "        And EXISTS (Select 1 FROM SAM_BENEFICIARIO X                           " + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "              WHERE CONTRATO = (Select CONTRATOMIGRACAO FROM SAM_SITUACAORH A  " + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                                 Join WEB_CADASTROACAO B On (B.SITUACAORH = A.HANDLE) " + Chr(13) + Chr(10)
        'vSelectEspecial = vSelectEspecial + "                                WHERE B.IDENTIFICADORVISAO = '"+WebVisionCode+"'))"


        vMensagem = "Não existe dependente para ser cancelado"
      End If

      BENEFICIARIO.WebLocalWhere = vSelectEspecial

      sql.Clear
      sql.Add("SELECT COUNT(A.HANDLE) QTD")
      sql.Add("  FROM SAM_BENEFICIARIO A")
      sql.Add("  JOIN SAM_FAMILIA      F ON (A.FAMILIA = F.HANDLE)")
      If pCodigoProcesso = 15 Then
        sql.Add("  JOIN SAM_MATRICULA  M ON (A.MATRICULA = M.Handle)")
      End If
      sql.Add("  WHERE ")
      sql.Add(vSelectEspecial)


      sql.Active = True

      If sql.FieldByName("QTD").AsInteger = 1 Then
        sql.Clear
        sql.Add("SELECT A.HANDLE")
        sql.Add("  FROM SAM_BENEFICIARIO A")
        sql.Add("  JOIN SAM_FAMILIA      F ON (A.FAMILIA = F.HANDLE)")
	    If pCodigoProcesso = 15 Then
	      sql.Add("  JOIN SAM_MATRICULA  M ON (A.MATRICULA = M.Handle)")
	    End If
        sql.Add("  WHERE ")
        sql.Add(vSelectEspecial)
        sql.Active = True

        CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = sql.FieldByName("HANDLE").AsInteger
      Else
        If sql.FieldByName("QTD").AsInteger = 0 Then
          InfoDescription = vMensagem
        End If

      End If


    End If

    On Error GoTo ERRO

    Set sql = Nothing
    Exit Function

    ERRO:
      WebVisionCode1 = False
      Set sql = Nothing
      InfoDescription = sql.Text

      CancelDescription = sql.Text 'Err.Description


End Function
