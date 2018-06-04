'HASH: 25453C7D2CC3ED3BBB8EB6B5C0BFCD1F
'#Uses "*bsShowMessage"
'#uses "*QueryToXML"
Option Explicit

Dim PRESTADOR As Long
Dim ESPECIALIDADE As Long
Dim DESTINATARIO As String
Dim ESPECIALIDADEDESTINO As Long

Public Sub TABLE_AfterInsert()

  CurrentQuery.FieldByName("PACIENTE").AsInteger = CLng(SessionVar("SAM_MATRICULA"))

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object

  If CurrentQuery.State = 3 Then

    Set sql = NewQuery

    GetAtendimento

    CurrentQuery.FieldByName("PRESTADOR").AsInteger = PRESTADOR
    CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger = ESPECIALIDADE
    CurrentQuery.FieldByName("DESTINATARIO").AsString = DESTINATARIO
    CurrentQuery.FieldByName("ESPECIALIDADEDESTINO").AsInteger = ESPECIALIDADEDESTINO

    Set sql = Nothing

  End If

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BUSCAREVENTOS" Then
    InfoDescription = BuscarEventos
  End If
    If CommandID = "BUSCARRESULTADOSEXAMES" Then
    InfoDescription = BuscarResultadosExames
  End If
  If CommandID = "BUSCARPROBLEMASPACIENTE" Then
    InfoDescription = BuscarProblemasPaciente
  End If
End Sub
Public Function BuscarEventos As String
  Dim sql As BPesquisa

  Set sql = NewQuery

  sql.Clear
  sql.Add(" SELECT T.HANDLE,                                               ")
  sql.Add("        T.ESTRUTURA,                                            ")
  sql.Add("        T.DESCRICAO,                                            ")
  sql.Add("        (SELECT ESTRUTURADO                                     ")
  sql.Add("           FROM CLI_MONITORAMENTO                               ")
  sql.Add("          WHERE PROCEDIMENTO = T.HANDLE) ESTRUTURADO            ")
  sql.Add("   FROM SAM_TGE    T                                            ")
  sql.Add("  WHERE T.CLASSEEVENTO IN (SELECT HANDLE FROM SAM_CLASSEEVENTO WHERE TIPO IN (3,4))")
  sql.Active = True

  BuscarEventos = QueryToXml(sql)

  Set sql = Nothing

End Function
Public Function BuscarProblemasPaciente As String
  Dim sql As BPesquisa

  Set sql = NewQuery

  sql.Clear
  sql.Add(" SELECT DISTINCT D.HANDLE, C2.ESTRUTURA, C2.DESCRICAO, D.EHCIDPRINCIPAL")
  sql.Add("   FROM CLI_EXAME_RESULTADO       R                                    ")
  sql.Add("   JOIN CLI_PLANO_EXAMES_EXAME    E ON E.HANDLE = R.EXAME              ")
  sql.Add("   JOIN CLI_PLANO_EXAMES_PROBLEMA P ON P.EXAMESOLICIT = E.EXAMESOLICIT ")
  sql.Add("   JOIN CLI_PACIENTEDIAGNOSTICO   D ON D.HANDLE = P.DIAGNOSTICO        ")
  sql.Add("   JOIN CLI_CID_CID               C1 ON C1.HANDLE = D.CID              ")
  sql.Add("   JOIN SAM_CID                   C2 ON C2.HANDLE = C1.CID             ")
  sql.Add("  WHERE R.DATA = :DATA                                                 ")
  sql.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATA").AsDateTime
  sql.Active = True

  BuscarProblemasPaciente = QueryToXml(sql)

  Set sql = Nothing

End Function
Public Function BuscarResultadosExames As String
  Dim sql As BPesquisa

  Set sql = NewQuery

  sql.Active = False
  sql.Clear
  sql.Add("SELECT T.HANDLE,                                                                ")
  sql.Add("       R.HANDLE HANDLERESULTADO,                                                ")
  sql.Add("       T.ESTRUTURA,                                                             ")
  sql.Add("       T.DESCRICAO,                                                             ")
  sql.Add("       R.RESULTADO,                                                             ")
  sql.Add("       (SELECT ESTRUTURADO                                                      ")
  sql.Add("          FROM CLI_MONITORAMENTO                                                ")
  sql.Add("         WHERE PROCEDIMENTO = E.EVENTO) ESTRUTURADO,                            ")
  sql.Add("       ESP1.DESCRICAO ESPECIALIDADE,                                            ")
  sql.Add("       ESP2.DESCRICAO ESPECIALIDADEDESTINO,                                     ")
  sql.Add("       P.NOME PRESTADOR,                                                        ")
  sql.Add("       R.DESTINATARIO,                                                          ")
  sql.Add("       R.PEDIDOEXTERNO                                                          ")
  sql.Add("  FROM CLI_EXAME_RESULTADO       R                                              ")
  sql.Add("  JOIN SAM_PRESTADOR             P ON P.HANDLE = R.PRESTADOR                    ")
  sql.Add("  JOIN SAM_ESPECIALIDADE         ESP1 ON ESP1.HANDLE = R.ESPECIALIDADE          ")
  sql.Add("  JOIN SAM_ESPECIALIDADE         ESP2 ON ESP2.HANDLE = R.ESPECIALIDADEDESTINO   ")
  sql.Add("  JOIN CLI_PLANO_EXAMES_EXAME    E ON E.HANDLE = R.EXAME                        ")
  sql.Add("  JOIN SAM_TGE                   T ON T.HANDLE = E.EVENTO                       ")
  sql.Add(" WHERE R.DATA = :DATA                                                           ")
  sql.Add("                                                                                ")
  sql.Add("UNION                                                                           ")
  sql.Add("                                                                                ")
  sql.Add("SELECT T.HANDLE,                                                                ")
  sql.Add("       R.HANDLE HANDLERESULTADO,                                                ")
  sql.Add("       T.ESTRUTURA,                                                             ")
  sql.Add("       T.DESCRICAO,                                                             ")
  sql.Add("       R.RESULTADO,                                                             ")
  sql.Add("       (SELECT ESTRUTURADO                                                      ")
  sql.Add("          FROM CLI_MONITORAMENTO                                                ")
  sql.Add("         WHERE PROCEDIMENTO = R.EVENTO) ESTRUTURADO,                            ")
  sql.Add("       ESP1.DESCRICAO ESPECIALIDADE,                                            ")
  sql.Add("       ESP2.DESCRICAO ESPECIALIDADEDESTINO,                                     ")
  sql.Add("       P.NOME PRESTADOR,                                                        ")
  sql.Add("       R.DESTINATARIO,                                                          ")
  sql.Add("       R.PEDIDOEXTERNO                                                          ")
  sql.Add("  FROM CLI_EXAME_RESULTADO       R                                              ")
  sql.Add("  JOIN SAM_PRESTADOR             P ON P.HANDLE = R.PRESTADOR                    ")
  sql.Add("  JOIN SAM_ESPECIALIDADE         ESP1 ON ESP1.HANDLE = R.ESPECIALIDADE          ")
  sql.Add("  JOIN SAM_ESPECIALIDADE         ESP2 ON ESP2.HANDLE = R.ESPECIALIDADEDESTINO   ")
  sql.Add("  JOIN SAM_TGE                   T ON T.HANDLE = R.EVENTO                       ")
  sql.Add(" WHERE R.DATA = :DATA                                                           ")
  sql.ParamByName("DATA").AsDateTime = CurrentQuery.FieldByName("DATA").AsDateTime
  sql.Active = True

  BuscarResultadosExames = QueryToXml(sql)

  Set sql = Nothing

End Function
Public Function GetAtendimento
  Dim sql As Object

  Set sql = NewQuery

  sql.Clear
  sql.Add(" SELECT R.PRESTADOR, S.ESPECIALIDADE, P.NOME             ")
  sql.Add("   FROM CLI_SUBJETIVO      S                             ")
  sql.Add("   JOIN CLI_RECURSO        R ON R.HANDLE = S.RECURSO     ")
  sql.Add("   JOIN SAM_PRESTADOR      P ON P.HANDLE = R.PRESTADOR   ")
  sql.Add("   JOIN CLI_USUARIO_SESSAO U ON U.ATENDIMENTO = S.HANDLE ")
  sql.Add("  WHERE U.USUARIO = :USUARIO                             ")
  sql.ParamByName("USUARIO").AsInteger = CurrentUser
  sql.Active = True

  PRESTADOR = sql.FieldByName("PRESTADOR").AsInteger
  ESPECIALIDADE = sql.FieldByName("ESPECIALIDADE").AsInteger
  DESTINATARIO = sql.FieldByName("NOME").AsString
  ESPECIALIDADEDESTINO = sql.FieldByName("ESPECIALIDADE").AsInteger

  Set sql = Nothing

End Function
