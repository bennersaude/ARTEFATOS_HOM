﻿'HASH: 1448E8E52765099B063CFA2B2C2B134B
 

Public Sub RELATORIOESPECIFICO_OnPopup(ShowPopup As Boolean)
  RELATORIOESPECIFICO.LocalWhere = "HANDLE IN (SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO IN ('BENCARTAO1', 'BENCARTAO2', 'BENCARTAO4', 'BENCARTAO5', 'BENCARTAO6', 'BENCARTAOCHESF', 'STJ034'))"
End Sub
