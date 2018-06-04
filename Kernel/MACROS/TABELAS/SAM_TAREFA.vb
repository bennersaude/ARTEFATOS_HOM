'HASH: 7E3D1990D0E098AF85FCC7943AFB244E
 

Public Sub DESCRICAO_OnExit()
  Dim interface As Object
  Set interface = CreateBennerObject("sampeg.Rotinas")

  Dim qAjustaTHM As Object
  Set qAjustaTHM = NewQuery

  Dim qBuscaDadosTHM As Object
  Set qBuscaDadosTHM = NewQuery

  qAjustaTHM.Clear
  qAjustaTHM.Add("UPDATE SAM_GUIA_EVENTOS SET XTHM = NULL WHERE HANDLE = :HANDLE")

  qBuscaDadosTHM.Clear
  qBuscaDadosTHM.Add("SELECT DISTINCT E.HANDLE                       ")
  qBuscaDadosTHM.Add("  FROM SAM_GUIA_EVENTOS E                      ")
  qBuscaDadosTHM.Add("  JOIN SAM_XTHM         T ON E.XTHM = T.HANDLE ")
  qBuscaDadosTHM.Add(" WHERE T.XTHM = 2                              ")
  qBuscaDadosTHM.Add("   AND E.FATURAPAGAMENTO IS NULL               ")
  qBuscaDadosTHM.Active = True

  While Not qBuscaDadosTHM.EOF

    qAjustaTHM.ParamByName("HANDLE").AsInteger = qBuscaDadosTHM.FieldByName("HANDLE").AsInteger
    qAjustaTHM.ExecSQL

    interface.RevisarEvento(CurrentSystem, qBuscaDadosTHM.FieldByName("HANDLE").AsInteger, "TOTAL", True)

    qBuscaDadosTHM.Next
  Wend

  Set qBuscaDadosTHM = Nothing
  Set qAjustaTHM = Nothing
  Set interface = Nothing
End Sub
