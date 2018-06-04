'HASH: 851E9D47B3A23E07389FE416376D9959
'#Uses "*bsShowMessage"
'#uses "*QueryToXML"

Option Explicit

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "BUSCAREVENTOSMONITORADOS" Then
      InfoDescription = BuscarEventosMonitorados
  End If
End Sub
Public Function BuscarEventosMonitorados As String
  Dim sql As BPesquisa

  Set sql = NewQuery

  sql.Clear
  sql.Add("(SELECT P.HANDLE, T.ESTRUTURA, T.DESCRICAO, M.MONITORAMENTOGERAL, 'S' PACIENTE ")
  sql.Add("   FROM CLI_PACIENTE_MONITORAMENTO P                                           ")
  sql.Add("   JOIN CLI_MONITORAMENTO          M ON M.HANDLE    = P.MONITORAMENTO          ")
  sql.Add("   JOIN SAM_TGE                    T ON T.HANDLE    = M.PROCEDIMENTO           ")
  sql.Add("   JOIN CLI_USUARIO_SESSAO         U ON U.MATRICULA = P.PACIENTE               ")
  sql.Add("  WHERE U.USUARIO = :USUARIO)                                                  ")
  sql.Add("UNION                                                                          ")
  sql.Add("(SELECT M.HANDLE, T.ESTRUTURA, T.DESCRICAO, M.MONITORAMENTOGERAL, 'N' PACIENTE ")
  sql.Add("   FROM CLI_MONITORAMENTO M                                                    ")
  sql.Add("   JOIN SAM_TGE           T ON T.HANDLE = M.PROCEDIMENTO)                      ")
  sql.ParamByName("USUARIO").AsInteger = CurrentUser
  sql.Active = True

  BuscarEventosMonitorados = QueryToXml(sql)

  Set sql = Nothing

End Function
