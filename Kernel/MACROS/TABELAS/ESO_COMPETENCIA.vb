'HASH: 259333FD9E1BEEB3C6AC1B98D8D780F6
'Tabela: ESO_COMPETENCIA

Public Sub RefreshCarga()
  RefreshNodesWithTable("ESO_COMPETENCIA")
End Sub

Public Sub BOTAOPROCESSAR_AfterOnClick()
  RefreshCarga
End Sub
