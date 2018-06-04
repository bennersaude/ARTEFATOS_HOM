'HASH: D20958A59754C5E48A4FBB24E5C8E7AA
'#Uses "*bsShowMessage"
Dim viPrestador As Long
Dim vsTextos As String



Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "CRIARUSUARIOSESSAO" Then
    viPrestador = -1
    InicializarVariaveisSessao(CurrentVirtualQuery.FieldByName("CLINICA").AsInteger, _
                               SelecionaRecurso(True), _
                               CurrentVirtualQuery.FieldByName("ESPECIALIDADE").AsInteger)
  End If

  If CommandID = "IDENTIFICARDADOSACESSO" Then
    InfoDescription = IdentificarDadosAcesso
  End If

  If CommandID = "LIMPARUSUARIOSESSAO" Then
    LimparUsuarioSessao
  End If
End Sub

Public Sub CriarUsuarioSessao()
  Dim SP As BStoredProc
  Set SP = NewStoredProc
  'Stored Procedure benner.sc.clinica.CriaUsuarioSessao
  SP.Name = "BS_3BB450BD"
  SP.AddParam("P_APELIDOUSUARIO",ptInput,sfString)
  SP.ParamByName("P_APELIDOUSUARIO").AsString = UserNickName
  SP.ExecProc
End Sub

Public Sub LimparUsuarioSessao()

    Dim SQL As Object
    Set SQL = NewQuery

    On Error GoTo Erro

    If Not InTransaction Then StartTransaction

    SQL.Clear
    SQL.Add(" UPDATE CLI_USUARIO_SESSAO             ")
    SQL.Add("    SET BENEFICIARIO = NULL,           ")
    SQL.Add("        PLANO = NULL,                  ")
    SQL.Add("        PRIORITARIO = NULL,            ")
    SQL.Add("        ESTRATIFICACAO = NULL,         ")
    SQL.Add("        EQUIPE = NULL,                 ")
    SQL.Add("        DATAATENDIMENTO = NULL,        ")
    SQL.Add("        CHAVE = NULL,                  ")
    SQL.Add("        CLINICA = NULL,                ")
    SQL.Add("        RECURSO = NULL,                ")
    SQL.Add("        MATRICULA = NULL,              ")
    SQL.Add("        ATENDIMENTO = NULL,            ")
    SQL.Add("        SAMPLANO = NULL,               ")
    SQL.Add("        ESPECIALIDADE = NULL           ")
    SQL.Add("  WHERE USUARIO = :USUARIO             ")
    SQL.ParamByName("USUARIO").AsInteger = CurrentUser
    SQL.ExecSQL

    GoTo Fim

    Erro:
    If InTransaction Then Rollback
    Exit Sub

    Fim:
    If InTransaction Then Commit


End Sub

Function InicializarVariaveisSessao(ByVal pHClinica As Long, ByVal pHRecurso As Long, ByVal pHEspecialidade As Long) As Boolean

  SessionVar("USUARIOCORRENTE")    = Str(CurrentUser)
  SessionVar("SAM_MATRICULA")      = Str(-1)
  SessionVar("CLI_SUBJETIVO")      = Str(-1)
  SessionVar("CLI_CLINICA")        = Str(pHClinica)
  SessionVar("SAM_ESPECIALIDADE")  = Str(pHEspecialidade)
  SessionVar("CLI_RECURSO")        = Str(pHRecurso)
  SessionVar("SAM_PLANO")          = Str(-1)
  SessionVar("SAM_BENEFICIARIO")   = Str(-1)

  SessionVar("DATACONSULTA")       = Str(ServerDate)
  SessionVar("PRESTADOR")          = Str(viPrestador) 'Linha temporária
  SessionVar("SAM_PRESTADOR")      = Str(viPrestador)

  Call InicializarSessao (pHClinica, pHRecurso, pHEspecialidade)

  InicializarVariaveisSessao = True

End Function

Public Function SelecionaRecurso(pbFiltrarClinica As Boolean) As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT R.HANDLE, P.HANDLE HPRESTADOR                       ")
  SQL.Add("  FROM CLI_RECURSO          R                              ")
  SQL.Add("  JOIN SAM_PRESTADOR        P ON (P.HANDLE = R.PRESTADOR)  ")
  SQL.Add("  JOIN CLI_RECURSO_USUARIO RU ON (RU.PRESTADOR = P.HANDLE) ")
  SQL.Add(" WHERE RU.USUARIO = :USUARIO                               ")
  SQL.Add("   AND R.DATAINICIAL <= :DATAATUAL                         ")
  SQL.Add("   AND (R.DATAFINAL > :DATAATUAL OR R.DATAFINAL IS NULL)     ")

  If pbFiltrarClinica Then
    SQL.Add("   AND R.CLINICA  = :CLINICA                             ")
  End If

  SQL.Add("   AND RU.PRONTUARIOMEDICO = 'S'                           ")
  SQL.ParamByName("USUARIO").AsInteger = CurrentUser

  If pbFiltrarClinica Then
    SQL.ParamByName("CLINICA").AsInteger = CLng(CurrentVirtualQuery.FieldByName("CLINICA").AsString)
  End If

  SQL.ParamByName("DATAATUAL").AsDateTime = ServerDate

  SQL.Active = True
  If Not SQL.EOF Then
    SelecionaRecurso = SQL.FieldByName("HANDLE").AsInteger
    viPrestador = SQL.FieldByName("HPRESTADOR").AsInteger
  Else
    SelecionaRecurso = -1
    viPrestador = -1
  End If

  SQL.Active = False
  Set SQL = Nothing
End Function

Public Function IdentificarDadosAcesso As String
'O retorno dessa função deve ser uma string com o formato HANDLECLINICA|HANDLERECURSO|HANDLEESPECIALIDADE
Dim HClinica As Long
Dim HRecurso As Long
Dim HEspecialidade As Long

HRecurso = SelecionaRecurso(False)
viPrestador = -1

If HRecurso = -1 Then 'Não é um médico
  HClinica       = -1
  HEspecialidade = -1
  InicializarVariaveisSessao(-1, -1, -1)
Else
  'O sistema irá identificar a quantas clínicas e especialidades este recurso está vinculado
  HClinica       = SelecionaClinica(HRecurso)
  HEspecialidade = SelecionaEspecialidade(HRecurso)
  If (HEspecialidade = 0) Then 'Se não tem especialidade, não será habilitada a interface
    InicializarVariaveisSessao(-1, -1, 0)
  ElseIf (HClinica > 0 And HRecurso > 0 And HEspecialidade > 0) Then
    InicializarVariaveisSessao(HClinica, HRecurso, HEspecialidade)
  End If
End If

IdentificarDadosAcesso = Trim(Str(HClinica)) + "|" + Trim(Str(HRecurso)) + "|" + Trim(Str(HEspecialidade))

End Function

Public Function SelecionaClinica(pHRecurso As Long) As Long
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(DISTINCT R.CLINICA) QTDE                ")
  SQL.Add("  FROM CLI_CLINICA   C                               ")
  SQL.Add("  JOIN CLI_RECURSO   R ON (R.CLINICA = C.HANDLE)     ")
  SQL.Add("  JOIN SAM_PRESTADOR S ON (S.HANDLE  = R.PRESTADOR)  ")
  SQL.Add("WHERE R.PRESTADOR = (SELECT PRESTADOR                ")
  SQL.Add("                       FROM CLI_RECURSO              ")
  SQL.Add("                      WHERE HANDLE = :HRECURSO)      ")
  SQL.Add("   AND R.DATAINICIAL <= :DATAATUAL                         ")
  SQL.Add("   AND (R.DATAFINAL > :DATAATUAL OR R.DATAFINAL IS NULL)     ")
  SQL.ParamByName("HRECURSO").AsInteger = pHRecurso
  SQL.ParamByName("DATAATUAL").AsDateTime = ServerDate
  SQL.Active = True

  'Se chegou aqui é porque está vinculado a pelo menos uma clínica, pois existe registro na CLI_RECURSO para o prestador
  If SQL.FieldByName("QTDE").AsInteger > 1 Then 'Existe mais de uma clínica, portanto o usuário vai ter que escolher
    SelecionaClinica = -1
  Else
    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT CLINICA            ")
    SQL.Add("  FROM CLI_RECURSO        ")
    SQL.Add(" WHERE HANDLE = :HRECURSO ")
    SQL.ParamByName("HRECURSO").AsInteger = pHRecurso
    SQL.Active = True
    SelecionaClinica = SQL.FieldByName("CLINICA").AsInteger
  End If

  SQL.Active = False
  Set SQL = Nothing
End Function
Public Sub InicializarSessao(pHClinica As Long, pHRecurso As Long, pHEspecialidade As Long)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(S.HANDLE) QTDE   ")
  SQL.Add("  FROM CLI_USUARIO_SESSAO S   ")
  SQL.Add("WHERE S.USUARIO = :HUSUARIO     ")
  SQL.ParamByName("HUSUARIO").AsInteger = CurrentUser
  SQL.Active = True

  If Not InTransaction Then StartTransaction

  On Error GoTo Erro

  If SQL.FieldByName("QTDE").Value = 0 Then
    SQL.Clear
    SQL.Add(" INSERT INTO CLI_USUARIO_SESSAO (HANDLE, USUARIO, CHAVE, CLINICA, RECURSO, ESPECIALIDADE) ")
    SQL.Add(" VALUES (:HANDLE, :USUARIO, :CHAVE, :CLINICA, :RECURSO, :ESPECIALIDADE )                  ")
    SQL.ParamByName("HANDLE").AsInteger = NewHandle("CLI_USUARIO_SESSAO")
    SQL.ParamByName("USUARIO").AsInteger = CurrentUser
    SQL.ParamByName("CHAVE").AsInteger = CurrentUser
    SQL.ParamByName("CLINICA").AsInteger = pHClinica
    SQL.ParamByName("RECURSO").AsInteger = pHRecurso
    SQL.ParamByName("ESPECIALIDADE").AsInteger = pHEspecialidade
    SQL.ExecSQL
  Else
    SQL.Clear
    SQL.Add(" UPDATE CLI_USUARIO_SESSAO      		")
    SQL.Add("    SET BENEFICIARIO = NULL,    		")
    SQL.Add("        PLANO = NULL,           		")
    SQL.Add("        PRIORITARIO = NULL,     		")
    SQL.Add("        ESTRATIFICACAO = NULL,  		")
    SQL.Add("        EQUIPE = NULL,          		")
    SQL.Add("        DATAATENDIMENTO = NULL,  		")
    SQL.Add("        CHAVE = :CHAVE,          		")
    SQL.Add("        CLINICA = :CLINICA,	 		")
    SQL.Add("        RECURSO = :RECURSO,  	 		")
    SQL.Add("        ESPECIALIDADE = :ESPECIALIDADE ")
    SQL.Add("  WHERE USUARIO = :USUARIO      ")
    SQL.ParamByName("USUARIO").AsInteger = CurrentUser
    SQL.ParamByName("CHAVE").AsInteger = CurrentUser
    SQL.ParamByName("CLINICA").AsInteger = pHClinica
    SQL.ParamByName("RECURSO").AsInteger = pHRecurso
    SQL.ParamByName("ESPECIALIDADE").AsInteger = pHEspecialidade
    SQL.ExecSQL
  End If

  Set SQL = Nothing

  GoTo Fim

  Erro:
  If InTransaction Then Rollback
  Exit Sub

  Fim:
  If InTransaction Then Commit

End Sub

Public Function SelecionaEspecialidade(pHRecurso As Long) As Long

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT COUNT(DISTINCT E.ESPECIALIDADE) QTDE                            ")
  SQL.Add("  FROM CLI_ESCALAPEP                E                                  ")
  SQL.Add("  JOIN CLI_RECURSO                  R ON (R.HANDLE  = E.RECURSO)       ")
  SQL.Add("  JOIN SAM_PRESTADOR                S ON (S.HANDLE  = R.PRESTADOR)     ")
  SQL.Add("  JOIN SAM_PRESTADOR_ESPECIALIDADE PE ON (PE.HANDLE = E.ESPECIALIDADE) ")
  SQL.Add("  JOIN SAM_ESPECIALIDADE           ES ON (ES.HANDLE = PE.ESPECIALIDADE)")
  SQL.Add(" WHERE R.PRESTADOR = (SELECT PRESTADOR                                 ")
  SQL.Add("                        FROM CLI_RECURSO                               ")
  SQL.Add("                       WHERE HANDLE = :HRECURSO)                       ")
  SQL.Add("   AND ES.SERVICOPROPRIO = 'S'                                         ")
  SQL.ParamByName("HRECURSO").AsInteger = pHRecurso
  SQL.Active = True

  If SQL.FieldByName("QTDE").AsInteger > 0 Then
    If SQL.FieldByName("QTDE").AsInteger > 1 Then 'Existe mais de uma especialidade, portanto o usuário vai ter que escolher
      SelecionaEspecialidade = -1
    Else
      SQL.Active = False
      SQL.Clear
      SQL.Add("SELECT DISTINCT PE.ESPECIALIDADE                                             ")
      SQL.Add("  FROM SAM_PRESTADOR_ESPECIALIDADE PE                                        ")
      SQL.Add("  JOIN CLI_ESCALAPEP                E ON (E.ESPECIALIDADE = PE.HANDLE)       ")
      SQL.Add("  JOIN SAM_PRESTADOR                P ON (P.HANDLE        = PE.PRESTADOR)    ")
      SQL.Add("  JOIN CLI_RECURSO                  R ON (R.PRESTADOR     = P.HANDLE)        ")
      SQL.Add("  JOIN SAM_ESPECIALIDADE           ES ON (ES.HANDLE       = PE.ESPECIALIDADE)")
      SQL.Add(" WHERE P.HANDLE IN (SELECT PRESTADOR                                         ")
      SQL.Add("                      FROM CLI_RECURSO                                       ")
      SQL.Add("                     WHERE HANDLE = :HRECURSO)                               ")
      SQL.Add("   AND ES.SERVICOPROPRIO = 'S'                                               ")
      SQL.ParamByName("HRECURSO").AsInteger = pHRecurso
      SQL.Active = True
      SelecionaEspecialidade = SQL.FieldByName("ESPECIALIDADE").AsInteger
    End If
  Else
    SelecionaEspecialidade = 0
  End If

  SQL.Active = False
  Set SQL = Nothing
End Function
