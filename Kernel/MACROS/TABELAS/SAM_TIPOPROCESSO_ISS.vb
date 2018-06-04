'HASH: 563EC8E08B3DF545469E27F55CDF03DE
Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		ISS.WebLocalWhere = "A.HANDLE IN (SELECT I.HANDLE                                     " + _
                   			"               FROM SFN_ISS I                                    " + _
                   			"              WHERE NOT EXISTS (SELECT 1                         " + _
                   			"                                  FROM SAM_TIPOPROCESSO_ISS TI   " + _
                   			"                                 WHERE TI.ISS          = I.HANDLE " + _
                   			"                                   AND TI.TIPOPROCESSO = @CAMPO(TIPOPROCESSO) ))"
    ElseIf VisibleMode Then
    	ISS.LocalWhere = "HANDLE IN (SELECT I.HANDLE                                     " + _
                   		 "             FROM SFN_ISS I                                    " + _
                		 "            WHERE NOT EXISTS (SELECT 1                         " + _
                   		 "                                FROM SAM_TIPOPROCESSO_ISS TI   " + _
                   		 "                               WHERE TI.ISS          = I.HANDLE" + _
                   		 "                                 AND TI.TIPOPROCESSO = @TIPOPROCESSO))"
    End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
If WebMode Then
		ISS.WebLocalWhere = "A.HANDLE IN (SELECT I.HANDLE                                     " + _
                   			"               FROM SFN_ISS I                                    " + _
                   			"              WHERE NOT EXISTS (SELECT 1                         " + _
                   			"                                  FROM SAM_TIPOPROCESSO_ISS TI   " + _
                   			"                                 WHERE TI.ISS          = I.HANDLE " + _
                   			"                                   AND TI.TIPOPROCESSO = @CAMPO(TIPOPROCESSO) ))"
    ElseIf VisibleMode Then
    	ISS.LocalWhere = "HANDLE IN (SELECT I.HANDLE                                     " + _
                   		 "             FROM SFN_ISS I                                    " + _
                		 "            WHERE NOT EXISTS (SELECT 1                         " + _
                   		 "                                FROM SAM_TIPOPROCESSO_ISS TI   " + _
                   		 "                               WHERE TI.ISS          = I.HANDLE" + _
                   		 "                                 AND TI.TIPOPROCESSO = @TIPOPROCESSO))"
    End If
End Sub

