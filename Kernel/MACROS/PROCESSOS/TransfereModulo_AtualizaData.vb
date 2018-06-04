'HASH: 562FB3E1B7C64B68E8ED106C69291E70

Public Sub Main

  On Error GoTo erro

	Dim psCriterios As String
	Dim psDataAdesao As String
	Dim psMensagem As String


	psDataAdesao = CStr( ServiceVar("psDataAdesao") )
	psMensagem = CStr(ServiceVar("psMensagem"))

	Dim vlHModBeneficiario As Long
	vlHModBeneficiario = CLng( SessionVar("HMODBENEFICIARIO") )

	Dim SQL As Object

	Set SQL = NewQuery


	SQL.Add("SELECT B.CONTRATO, B.FAMILIA								   ")
    SQL.Add("  FROM SAM_BENEFICIARIO B                                     ")
    SQL.Add("  JOIN SAM_BENEFICIARIO_MOD BM ON (B.HANDLE = BM.BENEFICIARIO)")
    SQL.Add(" WHERE BM.HANDLE = :HANDLE")
    SQL.ParamByName("HANDLE").AsInteger = vlHModBeneficiario

    SQL.Active = True

    Dim vsAux As String


    vsAux = CStr(SQL.FieldByName("CONTRATO").AsInteger)


    SQL.Active = False
    SQL.Clear

	SQL.Add("SELECT CM.PLANO, CM.MODULO                               ")
 	SQL.Add("  FROM SAM_CONTRATO_MOD CM		                          ")
 	SQL.Add("  JOIN SAM_BENEFICIARIO_MOD BM ON (CM.HANDLE = BM.MODULO)")
 	SQL.Add("WHERE BM.Handle = :HANDLE")
 	SQL.ParamByName("HANDLE").AsInteger = vlHModBeneficiario

 	SQL.Active = True

	Dim vlModOrigem As Long
	Dim vlPlanoOrigem As Long

	vlPlanoOrigem = SQL.FieldByName("PLANO").AsInteger
	vlModOrigem  = SQL.FieldByName("MODULO").AsInteger


 	psCriterios =  "SAM_CONTRATO_MOD.CONTRATO = " + vsAux + _
   						    " AND ((SAM_CONTRATO_MOD.MODULO <> " + CStr(vlModOrigem) + ")" + _
  							"  OR  (SAM_CONTRATO_MOD.MODULO = " + CStr(vlModOrigem) + _
                            " AND SAM_CONTRATO_MOD.PLANO <> " + CStr(vlPlanoOrigem) + "))"

    SQL.Active = False
    SQL.Clear

    SQL.Add("SELECT OBRIGATORIO , TIPOMODULO, COM.PLANO")
    SQL.Add("  FROM SAM_CONTRATO_MOD COM,SAM_MODULO MOD")
    SQL.Add(" WHERE COM.MODULO=MOD.HANDLE              ")
    SQL.Add("   AND COM.MODULO = " + CStr(vlModOrigem) +    "  And  COM.CONTRATO = " +vsAux)

	SQL.Active = True

    vsAux = SQL.FieldByName("OBRIGATORIO").AsString



 	psCriterios = psCriterios + " AND SAM_CONTRATO_MOD.OBRIGATORIO = '" + vsAux +  "'"

    If Not (SQL.FieldByName("TIPOMODULO").AsString = "S" ) Then
	  psCriterios = psCriterios + " AND SAM_MODULO.TIPOMODULO <> 'S' "
	End If
	  psCriterios = psCriterios + " AND DATACANCELAMENTO IS NULL "

    SQL.Active = False
	SQL.Clear

    SQL.Add("SELECT TIPOMODULO,     ")
	SQL.Add("       TIPOCOBERTURA   ")
	SQL.Add("  FROM SAM_MODULO      ")
	SQL.Add(" WHERE HANDLE = :MODULO")

    SQL.ParamByName("MODULO").AsInteger = vlModOrigem

    SQL.Active = True

  	psCriterios = psCriterios +  " AND SAM_MODULO.TIPOMODULO = '" + SQL.FieldByName("TIPOMODULO").AsString + "'"
  	psCriterios = psCriterios +  " AND SAM_MODULO.TIPOCOBERTURA = '"  + SQL.FieldByName("TIPOCOBERTURA").AsString  + "'"

	SQL.Active = False
	SQL.Clear

    SQL.Add("SELECT A.TITULARRESPONSAVEL,          ")
    SQL.Add("       A.TABRESPONSAVEL,              ")
    SQL.Add("       C.DATAINCLUSAOPREVIDENCIA      ")
	SQL.Add("  FROM SAM_FAMILIA A,                 ")
    SQL.Add("       SAM_BENEFICIARIO B,            ")
    SQL.Add("       SAM_MATRICULA C                ")
 	SQL.Add(" WHERE A.HANDLE = B.FAMILIA           ")
   	SQL.Add("   AND B.MATRICULA = C.HANDLE         ")
   	SQL.Add("   AND B.HANDLE = A.TITULARRESPONSAVEL")
   	SQL.Add("   AND A.HANDLE = :HFAMILIA")
   	SQL.Add("   AND C.DATAINCLUSAOPREVIDENCIA <= :DATAADESAO")

   	SQL.ParamByName("HFAMILIA").AsInteger = vsFamilia
    SQL.ParamByName("DATAADESAO").AsDateTime = CDate(psDataAdesao)

    SQL.Active = True

  	If  SQL.FieldByName("DATAINCLUSAOPREVIDENCIA").IsNull Then
    	psCriterios = psCriterios +   " AND SAM_CONTRATO_MOD.VERIFICARCONTRIBPREVIDENCIA =  'N' "
    End If

	ServiceVar("psCriterios") = psCriterios

	Set SQL = Nothing

	erro:
		psMensagem =  Err.Description
		ServiceVar( "psMensagem" ) = psMensagem

End Sub
