'HASH: F5852F31C37066BAAEE237D824938CF2

Option Explicit

Public Sub CODIGOPLANOOPERADORAVINCULACAO_OnPopup(ShowPopup As Boolean)
	CODIGOPLANOOPERADORAVINCULACAO.LocalWhere = " SAM_MODULO.HANDLE IN (SELECT MAX(M2.HANDLE)                                                         " + _
	                                            "                         FROM SAM_MODULO M2                                                          " + _
	                                            "                         JOIN SAM_CONTRATO_MOD CM2 ON CM2.MODULO = M2.HANDLE                         " + _
	                                            "                         JOIN SAM_REGISTROMS   SR2 ON SR2.HANDLE = CM2.REGISTROMS                    " + _
	                                            "                        WHERE M2.CODIGOPLANOOPERADORA IN (SELECT M3.CODIGOPLANOOPERADORA             " + _
	                                            " 			                                                 FROM SAM_MODULO M3                       " + _
	                                            "                                       		            WHERE M3.CODIGOPLANOOPERADORA IS NOT NULL " + _
												"                                                           GROUP BY M3.CODIGOPLANOOPERADORA)         " + _
												"                      	   AND CM2.OBRIGATORIO = 'S'                                                  " + _
												"                    	   AND SR2.ENVIADOSCPA = 'S'                                                  " + _
												"                        GROUP BY M2.CODIGOPLANOOPERADORA)                                            "
End Sub
