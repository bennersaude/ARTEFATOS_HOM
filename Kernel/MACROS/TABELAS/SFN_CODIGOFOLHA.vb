﻿'HASH: DCEB469742F2BEBC4663128B18921639

'MACRO SFN_CODIGOFOLHA


Public Sub TABLE_AfterScroll()
	If WebMode Then
		TIPOLANCAMENTOIMPORTACAO.WebLocalWhere = "A.HANDLE IN ( SELECT B.HANDLE FROM SFN_TIPOlANCFIN B JOIN SIS_TIPOLANCFIN TPFIN ON (TPFIN.HANDLE = B.TIPOLANCFIN) WHERE TPFIN.CODIGO IN (40,60,61,20,400))"
  	Else
  		TIPOLANCAMENTOIMPORTACAO.LocalWhere = "A.HANDLE IN ( SELECT B.HANDLE FROM SFN_TIPOLANCFIN B JOIN SIS_TIPOLANCFIN TPFIN ON (TPFIN.HANDLE = B.TIPOLANCFIN) WHERE TPFIN.CODIGO In (40,60,61,20,400))"
  	End If
End Sub