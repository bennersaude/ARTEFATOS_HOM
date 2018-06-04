'HASH: 81362394FF16D861836697658666A07B



Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		CATEGORIA.WebLocalWhere = "A.HANDLE IN (SELECT CP.HANDLE                                       	          " + _
                         		  "             FROM SAM_CATEGORIA_PRESTADOR CP                      			  " + _
                         		  "            WHERE NOT EXISTS (SELECT 1                           			  " + _
                         		  "                                FROM SAM_TIPOPROCESSO_CATEGORIA TC 		      " + _
                         		  "                               WHERE TC.CATEGORIA    = CP.HANDLE			      " + _
                         		  "                                 AND TC.TIPOPROCESSO = @CAMPO(TIPOPROCESSO) ))"
    ElseIf VisibleMode Then
    	CATEGORIA.LocalWhere = "HANDLE IN (SELECT A.HANDLE                                        	 " + _
                         	   "             FROM SAM_CATEGORIA_PRESTADOR A                      	 " + _
                         	   "            WHERE NOT EXISTS (SELECT 1                           	 " + _
                               "                                FROM SAM_TIPOPROCESSO_CATEGORIA B	 " + _
                         	   "                               WHERE B.CATEGORIA    = A.HANDLE    	 " + _
                         	   "                                 AND B.TIPOPROCESSO = @TIPOPROCESSO))"
    End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If WebMode Then
		CATEGORIA.WebLocalWhere = "A.HANDLE IN (SELECT CP.HANDLE                                       	          " + _
                         		  "             FROM SAM_CATEGORIA_PRESTADOR CP                      			  " + _
                         		  "            WHERE NOT EXISTS (SELECT 1                           			  " + _
                         		  "                                FROM SAM_TIPOPROCESSO_CATEGORIA TC 		      " + _
                         		  "                               WHERE TC.CATEGORIA    = CP.HANDLE			      " + _
                         		  "                                 AND TC.TIPOPROCESSO = @CAMPO(TIPOPROCESSO) ))"
    ElseIf VisibleMode Then
    	CATEGORIA.LocalWhere = "HANDLE IN (SELECT A.HANDLE                                        	 " + _
                         	   "             FROM SAM_CATEGORIA_PRESTADOR A                      	 " + _
                         	   "            WHERE NOT EXISTS (SELECT 1                           	 " + _
                               "                                FROM SAM_TIPOPROCESSO_CATEGORIA B	 " + _
                         	   "                               WHERE B.CATEGORIA    = A.HANDLE    	 " + _
                         	   "                                 AND B.TIPOPROCESSO = @TIPOPROCESSO))"
    End If
End Sub
